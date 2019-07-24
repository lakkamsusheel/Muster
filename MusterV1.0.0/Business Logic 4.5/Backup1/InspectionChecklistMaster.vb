'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.InspectionChecklistMaster
'   Provides the operations required to manipulate an Entity object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       MNR        06/11/05    Original class definition
'
' Function          Description
' GetEntity(NAME)   Returns the Entity requested by the string arg NAME
' GetEntity(ID)     Returns the Entity requested by the int arg ID
' GetAll()          Returns an ReportsCollection with all Entity objects
' Add(ID)           Adds the Entity identified by arg ID to the 
'                           internal ReportsCollection
' Add(Name)         Adds the Entity identified by arg NAME to the internal 
'                           ReportsCollection
' Add(Entity)       Adds the Entity passed as the argument to the internal 
'                           ReportsCollection
' Remove(ID)        Removes the Entity identified by arg ID from the internal 
'                           ReportsCollection
' Remove(NAME)      Removes the Entity identified by arg NAME from the 
'                           internal ReportsCollection
' EntityTable()     Returns a datatable containing all columns for the Entity 
'                           objects in the internal ReportsCollection.
'
' NOTE: This file to be used as InspectionChecklistMaster to build other objects.
'       Replace keyword "InspectionChecklistMaster" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'
'
'  6/2/2009   Thomas Franey   Added Business Rules for Checklist generator for 1.10, 2.11,2.12,4.4.5,4.8.4,
'                                                                              5.7.2, & 5.8.4
'
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pInspectionChecklistMaster
#Region "Public Events"
        Public Event evtInspectionChecklistMasterErr(ByVal MsgStr As String)
        Public Event evtInspectionChecklistMasterChanged(ByVal bolValue As Boolean)
        Public Event evtTankValidationErr(ByVal tnkID As Integer, ByVal strMessage As String)
#End Region
#Region "Private Member Variables"

        Private oInspection As MUSTER.Info.InspectionInfo
        Private WithEvents oInspectionChecklistMasterInfo As MUSTER.Info.InspectionChecklistMasterInfo
        Private oInspectionChecklistMasterDB As MUSTER.DataAccess.InspectionChecklistMasterDB
        Private MusterException As MUSTER.Exceptions.MusterExceptions


        Private oProperty As MUSTER.BusinessLogic.pProperty
        Private WithEvents oOwner As MUSTER.BusinessLogic.pOwner
        Private WithEvents oInspectionResponses As MUSTER.BusinessLogic.pInspectionResponse
        Private WithEvents oInspectionCPReadings As MUSTER.BusinessLogic.pInspectionCPReadings
        Private WithEvents oInspectionMonitorWells As MUSTER.BusinessLogic.pInspectionMonitorWells
        Private WithEvents oInspectionCCAT As MUSTER.BusinessLogic.pInspectionCCAT
        Private WithEvents oInspectionCitation As MUSTER.BusinessLogic.pInspectionCitation
        Private WithEvents oInspectionDiscrep As MUSTER.BusinessLogic.pInspectionDiscrep
        Private WithEvents oInspectionRectifier As MUSTER.BusinessLogic.pInspectionRectifier
        Private WithEvents oInspectionSketch As MUSTER.BusinessLogic.pInspectionSketch
        Private WithEvents oInspectionSOC As MUSTER.BusinessLogic.pInspectionSOC
        Private WithEvents oInspectionComments As MUSTER.BusinessLogic.pInspectionComments

        Private nID As Int64 = -1
        Private colIndex As Int64 = 0
        Private strErrMsg As String

        Public facpics As Collections.ArrayList
        Public facDocs As Collections.ArrayList

        Private htTankIDIndex As Hashtable
        Private htPipeIDIndex As Hashtable
        'Private colPipeIDIndex As Hashtable

        Private nMaxTankCPLineNum, nMaxPipeCPLineNum, nMaxTermCPLineNum, nMaxTankBedMWLineNum, nMaxLineMWLineNum, nMaxTankLineMWLineNum As Int64

        ' variables to generate checklist
        Public hasOwnerAsDesignatedOperator As Boolean

        Private hasCIU, hasTOS, hasTOSI, hasPipes, _
        CIUhasPipes, TOShasPipes, TOSIhasPipes, _
        CIUsingleWalledPost88NotInspected, _
        CIU_PPost88NotInspected, _
        CIU_PipeInstalledAfter10_1_08, _
        CIUfiberGlassPost88NotInspected, _
        CIU_Ppressurized, _
        CIUhazardousSubstance, _
        TOSIhazardousSubstance, _
        CIU_TanksInstalledAfter10_1_08, _
        CIUUsedOilOnly, _
        CIUballFloat, _
        CIUdropTube, _
        CIUelectronicAlarm, _
        CIUelectronicAlarmOnly, _
        CIUlikeCP, _
        CIUlikeGalvanic, _
        TOSlikeGalvanic, _
        CIU_PlikeCP, _
        TOSIlikeCP, _
        TOSI_PlikeCP, _
        CIUimpressedCurrent, _
        TOSIimpressedCurrent, _
        CIU_PimpressedCurrent, _
        TOSI_PimpressedCurrent, _
        CIU_PtermImpressedCurrent, _
        TOSI_PtermImpressedCurrent, _
        CIUcathodicallyProtected, _
        TOSIcathodicallyProtected, _
        CIULikecathodicallyProtected, _
        TOSILikecathodicallyProtected, _
        CIUlinedInteriorInstallgt10yrsAgo, _
        TOSIlinedInteriorInstallgt10yrsAgo, _
        CIU_PcathodicallyProtected, _
        TOSI_PcathodicallyProtected, _
        CIU_PLikecathodicallyProtected, _
        TOSI_PLikecathodicallyProtected, _
        CIU_Psteel, _
        TOSI_Psteel, _
        CIU_PtankContainedInBoots, _
        TOSI_PtankContainedInBoots, _
        CIU_PdispContainedInBoots, _
        TOSI_PdispContainedInBoots, _
        CIU_PtankContainedInSumps, _
        TOSI_PtankContainedInSumps, _
        CIU_PdispContainedInSumps, _
        TOSI_PdispContainedInSumps, _
        CIU_PdispCathodicallyProtected, _
        TOSI_PdispCathodicallyProtected, _
        CIU_PtankCathodicallyProtected, _
        TOSI_PtankCathodicallyProtected, _
        CIUemergenOnly, _
        CIUgroundWaterVaporMonitoring, _
        CIUinventoryControlPTT, _
        CIUautomaticTankGauging, _
        CIUstatisticalInventoryReconciliation, _
        CIUmanualTankGauging, _
        CIUvisualInterstitialMonitoring, _
        CIUelectronicInterstitialMonitoring, _
        CIU_PpressurizedUSSuction, _
        CIU_PgroundWaterVaporMonitoring, _
        CIU_PLTT, _
        CIU_PUSSuction, _
        CIU_PelectronicALLD, _
        CIU_Pmechanical, _
        CIU_PstatisticalInventoryReconciliation, _
        CIU_PvisualInterstitialMonitoring, _
        CIU_PcontinuousInterstitialMonitoring, _
        CIU_PnotContinuousInterstitialMonitoring, _
        CIU_Pelectronic, _
        CIU_Pplastic As Boolean
        Private AddCPReadingsTank, AddCPReadingsPipe, AddCPReadingsTerm As Boolean
        Private CPReadingsTankQID, CPReadingsPipeQID, CPReadingsTermQID As Int64
        Private AddMonitorWellsTank, AddMonitorWellsPipe As Boolean
        Private MonitorWellsTankQID, MonitorWellsPipeQID, MonitorWellsTankPipeQID As Int64

        Private qIDOfCLItem2point5 As Int64 = 0
        Private respOfCLItem2point5 As Int64 = -1
        Private qIDOfCLItem2point6 As Int64 = 0

        ' list to hold tankid's and pipeid's for cpreading
        ' so you can populate the drop down while adding 
        ' a new row in the UI for tank/pipe/term
        ' variables exposes through readonly property
        Private slTankID As New SortedList
        Private slPipeID As New SortedList
        Private slTermPipeID As New SortedList
        ' list to hold tank/pipe's fuel type for cp grid
        Private slTankFuelType As New SortedList
        Private slPipeFuelType As New SortedList
#End Region
#Region "Constructors"
        Public Sub New(Optional ByVal strDBConn As String = "", Optional ByRef MusterXCEP As MUSTER.Exceptions.MusterExceptions = Nothing, Optional ByRef inspection As MUSTER.Info.InspectionInfo = Nothing)
            If MusterXCEP Is Nothing Then
                MusterException = New MUSTER.Exceptions.MusterExceptions
            Else
                MusterException = MusterXCEP
            End If
            If inspection Is Nothing Then
                oInspection = New MUSTER.Info.InspectionInfo
            Else
                oInspection = inspection
            End If
            oInspectionChecklistMasterInfo = New MUSTER.Info.InspectionChecklistMasterInfo
            oInspectionChecklistMasterDB = New MUSTER.DataAccess.InspectionChecklistMasterDB

            htTankIDIndex = New Hashtable
            htPipeIDIndex = New Hashtable
            'colPipeIDIndex = New Hashtable

            oOwner = New MUSTER.BusinessLogic.pOwner(strDBConn, MusterXCEP)
            oInspectionResponses = New MUSTER.BusinessLogic.pInspectionResponse(strDBConn, MusterXCEP, oInspection)
            oInspectionCPReadings = New MUSTER.BusinessLogic.pInspectionCPReadings(strDBConn, MusterXCEP, oInspection)
            oInspectionMonitorWells = New MUSTER.BusinessLogic.pInspectionMonitorWells(strDBConn, MusterXCEP, oInspection)
            oInspectionCCAT = New MUSTER.BusinessLogic.pInspectionCCAT(strDBConn, MusterXCEP, oInspection)
            oInspectionCitation = New MUSTER.BusinessLogic.pInspectionCitation(strDBConn, MusterXCEP, oInspection)
            oInspectionDiscrep = New MUSTER.BusinessLogic.pInspectionDiscrep(strDBConn, MusterXCEP, oInspection)
            oInspectionRectifier = New MUSTER.BusinessLogic.pInspectionRectifier(strDBConn, MusterXCEP, oInspection)
            oInspectionSketch = New MUSTER.BusinessLogic.pInspectionSketch(strDBConn, MusterXCEP, oInspection)
            oInspectionSOC = New MUSTER.BusinessLogic.pInspectionSOC(strDBConn, MusterXCEP, oInspection)
            oInspectionComments = New MUSTER.BusinessLogic.pInspectionComments(strDBConn, MusterXCEP, oInspection)
        End Sub
#End Region
#Region "Exposed Attributes"
        Public ReadOnly Property ID() As Int64
            Get
                Return oInspectionChecklistMasterInfo.ID
            End Get
        End Property
        Public Property Position() As Int64
            Get
                Return oInspectionChecklistMasterInfo.Position
            End Get
            Set(ByVal Value As Int64)
                oInspectionChecklistMasterInfo.Position = Value
            End Set
        End Property
        Public Property CheckListItemNumber() As String
            Get
                Return oInspectionChecklistMasterInfo.CheckListItemNumber
            End Get
            Set(ByVal Value As String)
                oInspectionChecklistMasterInfo.CheckListItemNumber = Value
            End Set
        End Property
        Public Property SOC() As String
            Get
                Return oInspectionChecklistMasterInfo.SOC
            End Get
            Set(ByVal Value As String)
                oInspectionChecklistMasterInfo.SOC = Value
            End Set
        End Property
        Public Property Header() As Boolean
            Get
                Return oInspectionChecklistMasterInfo.Header
            End Get
            Set(ByVal Value As Boolean)
                oInspectionChecklistMasterInfo.Header = Value
            End Set
        End Property
        Public Property HeaderQuestionText() As String
            Get
                Return oInspectionChecklistMasterInfo.HeaderQuestionText
            End Get
            Set(ByVal Value As String)
                oInspectionChecklistMasterInfo.HeaderQuestionText = Value
            End Set
        End Property
        'Public Property ResponseTable() As String
        '    Get
        '        Return oInspectionChecklistMasterInfo.ResponseTable
        '    End Get
        '    Set(ByVal Value As String)
        '        oInspectionChecklistMasterInfo.ResponseTable = Value
        '    End Set
        'End Property
        Public ReadOnly Property AppliesToTank() As Boolean
            Get
                Return oInspectionChecklistMasterInfo.AppliesToTank
            End Get
        End Property
        Public ReadOnly Property AppliesToPipe() As Boolean
            Get
                Return oInspectionChecklistMasterInfo.AppliesToPipe
            End Get
        End Property
        Public ReadOnly Property AppliesToPipeTerm() As Boolean
            Get
                Return oInspectionChecklistMasterInfo.AppliesToPipeTerm
            End Get
        End Property
        Public Property Citation() As Int64
            Get
                Return oInspectionChecklistMasterInfo.Citation
            End Get
            Set(ByVal Value As Int64)
                oInspectionChecklistMasterInfo.Citation = Value
            End Set
        End Property
        Public Property DiscrepText() As String
            Get
                Return oInspectionChecklistMasterInfo.DiscrepText
            End Get
            Set(ByVal Value As String)
                oInspectionChecklistMasterInfo.DiscrepText = Value
            End Set
        End Property
        Public Property WhenVisible() As String
            Get
                Return oInspectionChecklistMasterInfo.WhenVisible
            End Get
            Set(ByVal Value As String)
                oInspectionChecklistMasterInfo.WhenVisible = Value
            End Set
        End Property
        Public Property CCAT() As Boolean
            Get
                Return oInspectionChecklistMasterInfo.CCAT
            End Get
            Set(ByVal Value As Boolean)
                oInspectionChecklistMasterInfo.CCAT = Value
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oInspectionChecklistMasterInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oInspectionChecklistMasterInfo.Deleted = Boolean.Parse(Value)
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oInspectionChecklistMasterInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oInspectionChecklistMasterInfo.IsDirty = Boolean.Parse(value)
            End Set
        End Property
        Public ReadOnly Property colIsDirty() As Boolean
            Get
                Return oInspectionCCAT.colIsDirty Or _
                        oInspectionCitation.colIsDirty Or _
                        oInspectionComments.colIsDirty Or _
                        oInspectionCPReadings.colIsDirty Or _
                        oInspectionDiscrep.colIsDirty Or _
                        oInspectionMonitorWells.colIsDirty Or _
                        oInspectionRectifier.colIsDirty Or _
                        oInspectionResponses.colIsDirty Or _
                        oInspectionSketch.colIsDirty Or _
                        oInspectionSOC.colIsDirty
            End Get
        End Property
        Public Property Show() As Boolean
            Get
                Return oInspectionChecklistMasterInfo.Show
            End Get
            Set(ByVal Value As Boolean)
                oInspectionChecklistMasterInfo.Show = Value
            End Set
        End Property
        Public Property CreatedBy() As String
            Get
                Return oInspectionChecklistMasterInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oInspectionChecklistMasterInfo.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oInspectionChecklistMasterInfo.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oInspectionChecklistMasterInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oInspectionChecklistMasterInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oInspectionChecklistMasterInfo.ModifiedOn
            End Get
        End Property
        Public Property Owner() As MUSTER.BusinessLogic.pOwner
            Get
                Return oOwner
            End Get
            Set(ByVal Value As MUSTER.BusinessLogic.pOwner)
                oOwner = Value
            End Set
        End Property
        Public Property InspectionInfo() As MUSTER.Info.InspectionInfo
            Get
                Return oInspection
            End Get
            Set(ByVal Value As MUSTER.Info.InspectionInfo)
                oInspection = Value
                SetInspectionToChild()
            End Set
        End Property
        Public Property InspectionResponses() As MUSTER.BusinessLogic.pInspectionResponse
            Get
                Return oInspectionResponses
            End Get
            Set(ByVal Value As MUSTER.BusinessLogic.pInspectionResponse)
                oInspectionResponses = Value
            End Set
        End Property
        Public Property InspectionCPReadings() As MUSTER.BusinessLogic.pInspectionCPReadings
            Get
                Return oInspectionCPReadings
            End Get
            Set(ByVal Value As MUSTER.BusinessLogic.pInspectionCPReadings)
                oInspectionCPReadings = Value
            End Set
        End Property
        Public Property InspectionMonitorWells() As MUSTER.BusinessLogic.pInspectionMonitorWells
            Get
                Return oInspectionMonitorWells
            End Get
            Set(ByVal Value As MUSTER.BusinessLogic.pInspectionMonitorWells)
                oInspectionMonitorWells = Value
            End Set
        End Property
        Public Property InspectionCCAT() As MUSTER.BusinessLogic.pInspectionCCAT
            Get
                Return oInspectionCCAT
            End Get
            Set(ByVal Value As MUSTER.BusinessLogic.pInspectionCCAT)
                oInspectionCCAT = Value
            End Set
        End Property
        Public Property InspectionCitation() As MUSTER.BusinessLogic.pInspectionCitation
            Get
                Return oInspectionCitation
            End Get
            Set(ByVal Value As MUSTER.BusinessLogic.pInspectionCitation)
                oInspectionCitation = Value
            End Set
        End Property
        Public Property InspectionDiscrep() As MUSTER.BusinessLogic.pInspectionDiscrep
            Get
                Return oInspectionDiscrep
            End Get
            Set(ByVal Value As MUSTER.BusinessLogic.pInspectionDiscrep)
                oInspectionDiscrep = Value
            End Set
        End Property
        Public Property InspectionRectifier() As MUSTER.BusinessLogic.pInspectionRectifier
            Get
                Return oInspectionRectifier
            End Get
            Set(ByVal Value As MUSTER.BusinessLogic.pInspectionRectifier)
                oInspectionRectifier = Value
            End Set
        End Property
        Public Property InspectionSketch() As MUSTER.BusinessLogic.pInspectionSketch
            Get
                Return oInspectionSketch
            End Get
            Set(ByVal Value As MUSTER.BusinessLogic.pInspectionSketch)
                oInspectionSketch = Value
            End Set
        End Property
        Public Property InspectionSOC() As MUSTER.BusinessLogic.pInspectionSOC
            Get
                Return oInspectionSOC
            End Get
            Set(ByVal Value As MUSTER.BusinessLogic.pInspectionSOC)
                oInspectionSOC = Value
            End Set
        End Property
        Public Property InspectionComments() As MUSTER.BusinessLogic.pInspectionComments
            Get
                Return oInspectionComments
            End Get
            Set(ByVal Value As MUSTER.BusinessLogic.pInspectionComments)
                oInspectionComments = Value
            End Set
        End Property
        Public ReadOnly Property CPReadingTankIDs() As SortedList
            Get
                Return slTankID
            End Get
        End Property
        Public ReadOnly Property CPReadingPipeIDs() As SortedList
            Get
                Return slPipeID
            End Get
        End Property
        Public ReadOnly Property CPReadingTermIDs() As SortedList
            Get
                Return slTermPipeID
            End Get
        End Property
        Public ReadOnly Property TankFuelType() As SortedList
            Get
                Return slTankFuelType
            End Get
        End Property
        Public ReadOnly Property PipeFuelType() As SortedList
            Get
                Return slPipeFuelType
            End Get
        End Property
        Public Property MaxTankCPLineNum() As Int64
            Get
                Return nMaxTankCPLineNum
            End Get
            Set(ByVal Value As Int64)
                nMaxTankCPLineNum = Value
            End Set
        End Property
        Public Property MaxPipeCPLineNum() As Int64
            Get
                Return nMaxPipeCPLineNum
            End Get
            Set(ByVal Value As Int64)
                nMaxPipeCPLineNum = Value
            End Set
        End Property
        Public Property MaxTermCPLineNum() As Int64
            Get
                Return nMaxTermCPLineNum
            End Get
            Set(ByVal Value As Int64)
                nMaxTermCPLineNum = Value
            End Set
        End Property
        Public Property MaxTankBedMWLineNum() As Int64
            Get
                Return nMaxTankBedMWLineNum
            End Get
            Set(ByVal Value As Int64)
                nMaxTankBedMWLineNum = Value
            End Set
        End Property
        Public Property MaxLineMWLineNum() As Int64
            Get
                Return nMaxLineMWLineNum
            End Get
            Set(ByVal Value As Int64)
                nMaxLineMWLineNum = Value
            End Set
        End Property
        Public Property MaxTankLineMWLineNum() As Int64
            Get
                Return nMaxTankLineMWLineNum
            End Get
            Set(ByVal Value As Int64)
                nMaxTankLineMWLineNum = Value
            End Set
        End Property
        Public ReadOnly Property QuestionIDOfCLItem2point5() As Int64
            Get
                Return qIDOfCLItem2point5
            End Get
        End Property
        Public Property ResponseOfCLItem2point5() As Int64
            Get
                Return respOfCLItem2point5
            End Get
            Set(ByVal Value As Int64)
                respOfCLItem2point5 = Value
            End Set
        End Property
        Public ReadOnly Property QuestionIDOfCLItem2point6() As Int64
            Get
                Return qIDOfCLItem2point6
            End Get
        End Property
        Public ReadOnly Property QuestionIDofCPItem354() As Int64
            Get
                Return CPReadingsTankQID
            End Get
        End Property
        Public ReadOnly Property QuestionIDofCPItem363() As Int64
            Get
                Return CPReadingsPipeQID
            End Get
        End Property
        Public ReadOnly Property QuestionIDofCPItem376() As Int64
            Get
                Return CPReadingsTermQID
            End Get
        End Property
        Public ReadOnly Property ShowMW() As Boolean
            Get
                Return Not (AddMonitorWellsTank And AddMonitorWellsPipe)
            End Get
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        Public Function RetrieveByCheckListItemNum(ByVal chkListItemNum As String) As MUSTER.Info.InspectionChecklistMasterInfo
            Try
                Return oInspectionChecklistMasterDB.DBGetByCheckListItemNum(chkListItemNum)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Sub RetrieveOwnerFacTanksPipes(ByRef inspection As MUSTER.Info.InspectionInfo)
            Try
                oInspection = inspection
                ' retrieve owner, facility, tanks, compartments, pipes
                If oOwner.ID <> oInspection.OwnerID Then
                    oOwner.Retrieve(oInspection.OwnerID)
                End If
                oOwner.Facilities.RetrieveAll(oInspection.OwnerID, "INSPECTION", False, oInspection.FacilityID, )
                ' set current facility to the facID supplied
                If oOwner.Facilities.ID <> oInspection.FacilityID Then
                    oOwner.Facilities.Retrieve(oOwner.OwnerInfo, oInspection.FacilityID, "SELF", "FACILITY", , True)
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub

        Public Function AllowsDBCheckListCode() As Boolean

            Try
                Return Me.oInspectionChecklistMasterDB.GetDBInspctionCheckListApproval
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex

            End Try
        End Function

        Public Function Retrieve(ByRef inspection As MUSTER.Info.InspectionInfo, ByVal inspectionID As Int64, ByVal facID As Int64, ByVal ownerID As Int64, Optional ByVal [readOnly] As Boolean = False, Optional ByVal ownerIsDesignatedOperator As Boolean = True, Optional ByVal path As String = "") As MUSTER.Info.InspectionChecklistMasterInfo
            Try
                oInspection = inspection


                ' retrieve owner, facility, tanks, compartments, pipes
                If oOwner.ID <> ownerID Then
                    If oInspection.SubmittedDate.ToString Is DBNull.Value Or oInspection.SubmittedDate = Date.MinValue Or oInspection.SubmittedDate.ToString = String.Empty Then
                        oOwner.Retrieve(ownerID)
                    Else
                        oOwner.Retrieve(ownerID, , , , oInspection.ID)
                    End If
                End If

                If (oInspection.SubmittedDate.ToString Is DBNull.Value) Or (oInspection.SubmittedDate = Date.MinValue) Or (oInspection.SubmittedDate.ToString = String.Empty) Then
                    oOwner.Facilities.RetrieveAll(ownerID, "INSPECTION", False, facID, , )
                Else
                    oOwner.Facilities.RetrieveAll(ownerID, "INSPECTION", False, facID, , oInspection.ID)
                End If
                ' set current facility to the facID supplied
                If oOwner.Facilities.ID <> facID Then
                    oOwner.Facilities.Retrieve(oOwner.OwnerInfo, facID, "SELF", "FACILITY", , True)
                End If



                ' retrieve all questions if not already retrieved
                If oInspection.ChecklistMasterCollection.Count = 0 Or Me.oInspectionChecklistMasterDB.GetDBInspctionCheckListApproval Then
                    RetrieveAllChecklistText(facID)
                End If
                ' retrieve checklist responses if collection is empty
                If oInspection.ResponsesCollection.Count = 0 Then
                    ' retrieveresponses - return true if there are values in the db, false if not
                    If RetrieveResponses(inspectionID, [readOnly], path) Then
                        GenerateCheckList(True, [readOnly], ownerIsDesignatedOperator)
                    Else
                        ' if no values are present in the db, it is the first time the checklist is being generated
                        ' set inspection's checklist generated date to now
                        oInspection.CheckListGenDate = Now.Date
                        GenerateCheckList(False, [readOnly], ownerIsDesignatedOperator)
                        ' create a row in "tblINS_INSPECTION_DATES" table for current inspection as there is no entry in the table for current inspection
                        ' if there is no date in rescheduledate field, use scheduled date field else use rescheduled field
                        'If Date.Compare(oInspection.RescheduledDate, CDate("01/01/0001")) = 0 Then
                        '    PutCLInspectionHistory(0, oInspection.ID, 0, oInspection.ScheduledDate, oInspection.ScheduledTime, String.Empty, False)
                        'Else
                        '    PutCLInspectionHistory(0, oInspection.ID, 0, oInspection.RescheduledDate, oInspection.RescheduledTime, String.Empty, False)
                        'End If
                    End If
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Sub RetrieveAllChecklistText(Optional ByVal fac_id As Integer = 0)
            Try
                oInspection.ChecklistMasterCollection = oInspectionChecklistMasterDB.DBGetChecklistText(Nothing, fac_id, oInspectionChecklistMasterDB.GetDBInspctionCheckListApproval, oInspection.ID)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Private Function RetrieveResponses(ByVal inspectionID As Int64, Optional ByVal [readOnly] As Boolean = False, Optional ByVal Path As String = "") As Boolean
            Try
                Dim ds As DataSet
                ds = oInspectionChecklistMasterDB.DBGetResponses(inspectionID)
                ' 0 - Responses
                ' 1 - CPReadings
                ' 2 - MonitorWells
                ' 3 - CCAT
                ' 4 - Citation
                ' 5 - Discrep
                ' 6 - Rectifier
                ' 7 - Sketch
                ' 8 - SOC
                ' 9 - Comments
                ds.Tables(0).TableName = "Responses"
                ds.Tables(1).TableName = "CPReadings"
                ds.Tables(2).TableName = "MonitorWells"
                ds.Tables(3).TableName = "CCAT"
                ds.Tables(4).TableName = "Citation"
                ds.Tables(5).TableName = "Discrep"
                ds.Tables(6).TableName = "Rectifier"
                ds.Tables(7).TableName = "Sketch"
                ds.Tables(8).TableName = "SOC"
                ds.Tables(9).TableName = "Comments"

                If ds.Tables("Responses").Rows.Count > 0 Then
                    oInspectionResponses.Load(oInspection, ds)
                End If
                If ds.Tables("CPReadings").Rows.Count > 0 Then
                    oInspectionCPReadings.Load(oInspection, ds)
                End If
                If ds.Tables("MonitorWells").Rows.Count > 0 Then
                    oInspectionMonitorWells.Load(oInspection, ds)
                End If
                If ds.Tables("CCAT").Rows.Count > 0 Then
                    oInspectionCCAT.Load(oInspection, ds)
                End If
                If ds.Tables("Citation").Rows.Count > 0 Then
                    oInspectionCitation.Load(oInspection, ds)
                End If
                If ds.Tables("Discrep").Rows.Count > 0 Then
                    oInspectionDiscrep.Load(oInspection, ds)
                End If
                If ds.Tables("Rectifier").Rows.Count > 0 Then
                    oInspectionRectifier.Load(oInspection, ds)
                End If
                If ds.Tables("Sketch").Rows.Count > 0 Then
                    oInspectionSketch.Load(oInspection, ds)
                End If
                If ds.Tables("SOC").Rows.Count > 0 Then
                    oInspectionSOC.Load(oInspection, ds)
                End If
                If ds.Tables("Comments").Rows.Count > 0 Then
                    oInspectionComments.Load(oInspection, ds)
                End If

                LoadFacilityPics(Path, oInspection.ID)

                Return IIf(oInspection.ResponsesCollection.Count > 0, True, False)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Sub LoadFacilityPics(ByVal path As String, Optional ByVal inspection_ID As Integer = 0)


            Dim dirpath As IO.DirectoryInfo
            Dim currdate As DateTime
            Dim prevdate As DateTime


            Try

                If facpics Is Nothing Then
                    facpics = New Collections.ArrayList
                End If

                facpics.Clear()

                If Not oInspectionSketch Is Nothing Then
                    oInspectionSketch.Pics.Clear()
                    oInspectionSketch.Pics = facpics
                End If

                oInspectionChecklistMasterDB.DBSetDateRangeOnInspection(inspection_ID, oInspection.FacilityID, prevdate, currdate)

                dirpath = New IO.DirectoryInfo(String.Format("{0}", path))
                For Each f As IO.FileInfo In dirpath.GetFiles(String.Format("*{0}*.jpg", oInspection.FacilityID))

                    Dim PicDate As DateTime = f.LastWriteTime
                    Dim name As String = f.Name.ToUpper

                    If name.Substring(name.IndexOf("_") + 1).StartsWith("F_") AndAlso name.StartsWith("INSPECTION_") Then
                        name = name.Substring(name.IndexOf("_") + 1)
                    End If

                    If name.ToUpper.StartsWith("F_") Then
                        Dim pullstr As String = name.Replace(String.Format("F_{0}_", oInspection.FacilityID), String.Empty).Substring(0, 10)
                        If IsDate(pullstr) Then
                            PicDate = Convert.ToDateTime(pullstr)
                        End If

                    End If


                    If Date.Compare(PicDate, currdate) <= 0 AndAlso Date.Compare(PicDate, prevdate) > 0 Then
                        With name.Replace("_", "")
                            Dim add As Integer = 0

                            If name.StartsWith("F_") OrElse name.StartsWith("_") Then
                                add += 1
                            End If

                            If .StartsWith(String.Format("F{0}", oInspection.FacilityID)) AndAlso .Length > (oInspection.FacilityID.ToString.Length + 1 + add) AndAlso Not IsNumeric(name.Substring(oInspection.FacilityID.ToString.Length + 1 + add, 1)) Then
                                Me.facpics.Add(f)
                            ElseIf .StartsWith(String.Format("{0}", oInspection.FacilityID)) AndAlso .Length > (oInspection.FacilityID.ToString.Length + add) AndAlso Not IsNumeric(name.Substring(oInspection.FacilityID.ToString.Length + add, 1)) Then
                                Me.facpics.Add(f)
                            End If
                        End With
                    End If
                Next


            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                dirpath = Nothing
            End Try


        End Sub

        Public Function GetThisInspectionDate() As DateTime
            Dim prevDate, currDate As DateTime

            oInspectionChecklistMasterDB.DBSetDateRangeOnInspection(oInspection.ID, oInspection.FacilityID, prevDate, currDate)

            Return currDate

        End Function

        Public Sub LoadFacilityScannedDocs(ByVal path As String, Optional ByVal inspection_ID As Integer = 0)


            Dim dirpath As IO.DirectoryInfo
            Dim currdate As DateTime
            Dim prevdate As DateTime


            Try

                If facDocs Is Nothing Then
                    facDocs = New Collections.ArrayList
                End If

                facDocs.Clear()


                oInspectionChecklistMasterDB.DBSetDateRangeOnInspection(inspection_ID, oInspection.FacilityID, prevdate, currdate)

                dirpath = New IO.DirectoryInfo(String.Format("{0}", path))
                For Each f As IO.FileInfo In dirpath.GetFiles(String.Format("*{0}*.pdf", oInspection.FacilityID))

                    If Date.Compare(f.LastWriteTime, currdate) < 0 AndAlso Date.Compare(f.LastWriteTime, prevdate) > 0 Then
                        With f.Name.Replace("_", "").ToUpper
                            If .StartsWith(String.Format("F{0}", oInspection.FacilityID)) AndAlso .Length > (oInspection.FacilityID.ToString.Length + 1) AndAlso Not IsNumeric(f.Name.Substring(oInspection.FacilityID.ToString.Length + 1, 1)) Then
                                Me.facDocs.Add(f)
                            ElseIf .StartsWith(String.Format("{0}", oInspection.FacilityID)) AndAlso .Length > oInspection.FacilityID.ToString.Length AndAlso Not IsNumeric(f.Name.Substring(oInspection.FacilityID.ToString.Length, 1)) Then
                                Me.facDocs.Add(f)
                            End If
                        End With
                    End If
                Next


            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                dirpath = Nothing
            End Try


        End Sub

        Public Sub GenerateCheckList(ByVal hasResponses As Boolean, Optional ByVal [readOnly] As Boolean = False, Optional ByVal OwnerIsDesignatedOperator As Boolean = True)
            Dim oInspCLInfo As MUSTER.Info.InspectionChecklistMasterInfo
            Try

                Me.hasOwnerAsDesignatedOperator = OwnerIsDesignatedOperator


                ' reset checklist show to false
                ResetCLShow()
                ' mark all responses deleted
                ' only the responses that are shown will be unmarked deleted
                If Not [readOnly] Then
                    MarkResponsesDeleted(True)
                End If
                InitCLItem2point5_6Variables()
                InitCLVariables()
                ' populate variables according to conditions in DDD
                CheckCLVariables()


                If Not Me.oInspectionChecklistMasterDB.GetDBInspctionCheckListApproval Then
                    For Each oInspCLInfo In oInspection.ChecklistMasterCollection.Values
                        ' If oInspCLInfo.ID > -20 Then

                        If (oInspCLInfo.CheckListItemNumber.Trim).StartsWith("99.") Then
                            ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                        Else
                            Select Case (oInspCLInfo.CheckListItemNumber.Trim)
                                Case "12"
                                    ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                Case "1.13"
                                    If hasOwnerAsDesignatedOperator Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If

                                Case "1"
                                    ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                Case "1.1"
                                    If hasCIU Or hasTOSI Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "1.2"
                                    If hasCIU OrElse hasTOSI Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "1.3"
                                    If hasCIU OrElse hasTOSI Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "1.4"
                                    If (hasCIU OrElse hasTOSI) AndAlso (CIUhasPipes OrElse TOSIhasPipes) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "1.5"
                                    If hasCIU AndAlso CIUsingleWalledPost88NotInspected Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "1.6"
                                    ' If hasCIU AndAlso CIU_PPost88NotInspected Then
                                     If hasCIU 
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                     End If
                                Case "1.7"
                                    If hasCIU AndAlso CIUfiberGlassPost88NotInspected Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "1.8"
                                    If hasCIU AndAlso CIU_Ppressurized Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "1.9"
                                    If (hasCIU OrElse hasTOSI) AndAlso (CIUhazardousSubstance OrElse TOSIhazardousSubstance) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "1.10"
                                    If hasCIU AndAlso CIU_Ppressurized Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "1.10.1"
                                    If hasCIU AndAlso CIU_Ppressurized Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "1.10.2"
                                    If hasCIU AndAlso CIU_Ppressurized Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If

                                Case "1.11"
                                    If hasCIU AndAlso CIU_TanksInstalledAfter10_1_08 Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "1.12"
                                    If hasCIU AndAlso CIU_PipeInstalledAfter10_1_08 Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "1.16"
                                    If hasCIU AndAlso CIU_PipeInstalledAfter10_1_08 Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "1.17"
                                    If hasCIU AndAlso CIU_PipeInstalledAfter10_1_08 Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "1.16.1"
                                    If hasCIU AndAlso CIU_PipeInstalledAfter10_1_08 Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If

                                Case "1.17.1"
                                    If hasCIU AndAlso CIU_PipeInstalledAfter10_1_08 Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If

                                Case "1.15"
                                    If hasCIU AndAlso Me.CIU_TanksInstalledAfter10_1_08 Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "1.15.1"
                                    If hasCIU AndAlso Me.CIU_TanksInstalledAfter10_1_08 Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "1.18.1"
                                    If hasCIU OrElse hasTOS OrElse hasTOSI Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "12.1"
                                    If hasCIU OrElse hasTOS OrElse hasTOSI Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "12.2"
                                    If (hasCIU OrElse hasTOS OrElse hasTOSI) And oInspCLInfo.ID < 9999 Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If

                                Case "12.3"  'Records must be maintained that document the groundwater/vapor monitoring wells have been checked once every month, is this true?
                                    If (hasCIU OrElse hasTOS OrElse hasTOSI) And oInspCLInfo.ID < 9999 Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If

                                Case "12.4"
                                    If (hasCIU OrElse hasTOS OrElse hasTOSI) And oInspCLInfo.ID < 9999 Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If


                                Case "2"
                                    If hasCIU AndAlso Not CIUUsedOilOnly Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    ElseIf hasCIU Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    ElseIf hasCIU AndAlso CIUballFloat Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    ElseIf hasCIU AndAlso CIUdropTube Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    ElseIf hasCIU AndAlso CIUelectronicAlarm Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    ElseIf hasCIU AndAlso Not (CIUUsedOilOnly Or CIUelectronicAlarmOnly) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "2.1"
                                    If hasCIU Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "2.2"
                                    If hasCIU AndAlso Not CIUUsedOilOnly Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "2.3"
                                    If hasCIU AndAlso Not CIUUsedOilOnly Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "2.3.1"
                                    If hasCIU AndAlso Not CIUUsedOilOnly Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If

                                Case "2.4"
                                    If hasCIU AndAlso Not CIUUsedOilOnly Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "2.5"
                                    If hasCIU AndAlso CIUballFloat Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                        qIDOfCLItem2point5 = oInspCLInfo.ID
                                    End If
                                Case "2.6"
                                    If hasCIU AndAlso CIUballFloat Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                        qIDOfCLItem2point6 = oInspCLInfo.ID
                                    End If
                                Case "2.7"
                                    If hasCIU AndAlso CIUdropTube Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "2.8"
                                    If hasCIU AndAlso CIUelectronicAlarm Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "2.9"
                                    If hasCIU AndAlso CIUballFloat Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "2.10"
                                    If hasCIU AndAlso Not (CIUUsedOilOnly Or CIUelectronicAlarmOnly) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If


                                Case "2.11"
                                    If hasCIU AndAlso Not (CIUUsedOilOnly) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "2.11.1"
                                    If hasCIU AndAlso Not (CIUUsedOilOnly) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If

                                Case "2.12"
                                    If hasCIU AndAlso Not (CIUUsedOilOnly) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "2.12.1"
                                    If hasCIU AndAlso Not (CIUUsedOilOnly) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If

                                Case "2.13"
                                    If hasCIU AndAlso Not (CIUUsedOilOnly) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If

                                Case "2.14"
                                    If hasCIU AndAlso Not (CIUUsedOilOnly) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If

                                Case "2.15"
                                    If hasCIU AndAlso Not (CIUUsedOilOnly) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If

                                Case "2.16"
                                    If hasCIU AndAlso CIUballFloat Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                        qIDOfCLItem2point6 = oInspCLInfo.ID
                                    End If

                                Case "3"
                                    If hasCIU OrElse hasTOSI Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    ElseIf (hasCIU OrElse hasTOSI) AndAlso (CIUlikeCP OrElse TOSIlikeCP OrElse CIU_PlikeCP OrElse TOSI_PlikeCP) Then
                                    ElseIf (hasCIU OrElse hasTOSI) AndAlso (CIUlikeCP OrElse TOSIlikeCP OrElse CIU_PlikeCP OrElse TOSI_PlikeCP) AndAlso (CIUimpressedCurrent OrElse TOSIimpressedCurrent OrElse CIU_PimpressedCurrent OrElse TOSI_PimpressedCurrent OrElse CIU_PtermImpressedCurrent OrElse TOSI_PtermImpressedCurrent) Then
                                    ElseIf (hasCIU OrElse hasTOSI) AndAlso (CIUlikeCP OrElse TOSIlikeCP OrElse CIU_PlikeCP OrElse TOSI_PlikeCP) AndAlso (CIUimpressedCurrent OrElse TOSIimpressedCurrent OrElse CIU_PimpressedCurrent OrElse TOSI_PimpressedCurrent) Then
                                    End If
                                Case "3.1"
                                    If (hasCIU OrElse hasTOSI) AndAlso (CIUlikeCP OrElse TOSIlikeCP OrElse CIU_PlikeCP OrElse TOSI_PlikeCP) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "3.2"
                                    If (hasCIU OrElse hasTOSI) AndAlso (CIUlikeCP OrElse TOSIlikeCP OrElse CIU_PlikeCP OrElse TOSI_PlikeCP) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "3.2.1"
                                    If (hasCIU OrElse hasTOSI) AndAlso (CIUlikeGalvanic OrElse TOSlikeGalvanic) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "3.3"
                                    If (hasCIU OrElse hasTOSI) AndAlso (CIUlikeCP OrElse TOSIlikeCP OrElse CIU_PlikeCP OrElse TOSI_PlikeCP) AndAlso (CIUimpressedCurrent OrElse TOSIimpressedCurrent OrElse CIU_PimpressedCurrent OrElse TOSI_PimpressedCurrent OrElse CIU_PtermImpressedCurrent OrElse TOSI_PtermImpressedCurrent) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "3.4"
                                    If (hasCIU OrElse hasTOSI) AndAlso (CIUlikeCP OrElse TOSIlikeCP OrElse CIU_PlikeCP OrElse TOSI_PlikeCP) AndAlso (CIUimpressedCurrent OrElse TOSIimpressedCurrent OrElse CIU_PimpressedCurrent OrElse TOSI_PimpressedCurrent) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                        AddRectifier(hasResponses, oInspCLInfo.ID, [readOnly])
                                    End If
                                Case "3.5"
                                    ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                Case "3.5.1"
                                    If hasCIU OrElse hasTOSI Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "3.5.2"
                                    If (hasCIU OrElse hasTOSI) AndAlso (CIULikecathodicallyProtected OrElse TOSILikecathodicallyProtected) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "3.5.3"
                                    If (hasCIU OrElse hasTOSI) AndAlso (CIUlinedInteriorInstallgt10yrsAgo OrElse TOSIlinedInteriorInstallgt10yrsAgo) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "3.5.4"
                                    If (hasCIU OrElse hasTOSI) AndAlso (CIULikecathodicallyProtected OrElse TOSILikecathodicallyProtected) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                        ' TODO - determine galvanic / impressed current
                                        AddCPReadingsTank = True
                                        CPReadingsTankQID = oInspCLInfo.ID
                                    End If
                                    'Case "3.5.4.1"
                                    '    If (hasCIU Orelse hasTOSI) Andalso (CIUcathodicallyProtected Orelse TOSIcathodicallyProtected) Then
                                    '        CPReadingsTankCitationID = oInspCLInfo.Citation
                                    '        CPReadingsTankCitationDesc = oInspCLInfo.DiscrepText
                                    '    End If
                                Case "3.6"
                                    If (hasCIU OrElse hasTOSI) AndAlso (CIUhasPipes OrElse TOSIhasPipes) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    ElseIf (hasCIU OrElse hasTOSI) AndAlso (CIU_PLikecathodicallyProtected OrElse TOSI_PLikecathodicallyProtected) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "3.6.1"
                                    If (hasCIU OrElse hasTOSI) AndAlso (CIUhasPipes OrElse TOSIhasPipes) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "3.6.2"
                                    If (hasCIU OrElse hasTOSI) AndAlso (CIU_PLikecathodicallyProtected OrElse TOSI_PLikecathodicallyProtected) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "3.6.3"
                                    If (hasCIU OrElse hasTOSI) AndAlso (CIU_PLikecathodicallyProtected OrElse TOSI_PLikecathodicallyProtected) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                        AddCPReadingsPipe = True
                                        CPReadingsPipeQID = oInspCLInfo.ID
                                    End If
                                    'Case "3.6.3.1"
                                    '    If (hasCIU Orelse hasTOSI) Andalso (CIU_PcathodicallyProtected Orelse TOSI_PcathodicallyProtected) Then
                                    '        CPReadingsPipeCitationID = oInspCLInfo.Citation
                                    '        CPReadingsPipeCitationDesc = oInspCLInfo.DiscrepText
                                    '    End If
                                Case "3.7"
                                    If (hasCIU OrElse hasTOSI) AndAlso Not ((CIUhasPipes AndAlso TOSIhasPipes AndAlso CIU_Psteel AndAlso TOSI_Psteel) OrElse _
                                                                    (CIUhasPipes AndAlso Not TOSIhasPipes AndAlso CIU_Psteel AndAlso Not TOSI_Psteel) OrElse _
                                                                    (Not CIUhasPipes AndAlso Not TOSIhasPipes AndAlso Not CIU_Psteel AndAlso Not TOSI_Psteel) OrElse _
                                                                    (Not CIUhasPipes AndAlso TOSIhasPipes AndAlso Not CIU_Psteel AndAlso TOSI_Psteel)) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    ElseIf (hasCIU OrElse hasTOSI) AndAlso (CIU_PtankContainedInBoots OrElse TOSI_PtankContainedInBoots OrElse CIU_PdispContainedInBoots OrElse TOSI_PdispContainedInBoots) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    ElseIf (hasCIU OrElse hasTOSI) AndAlso (CIU_PdispCathodicallyProtected OrElse TOSI_PdispCathodicallyProtected OrElse CIU_PtankCathodicallyProtected OrElse TOSI_PtankCathodicallyProtected) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "3.7.1"
                                    If (hasCIU OrElse hasTOSI) AndAlso Not ((CIUhasPipes AndAlso CIU_Psteel) OrElse (TOSIhasPipes AndAlso TOSI_Psteel)) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "3.7.2"
                                    If (hasCIU OrElse hasTOSI) AndAlso Not ((CIUhasPipes AndAlso CIU_Psteel) OrElse (TOSIhasPipes AndAlso TOSI_Psteel)) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "3.7.3"
                                    If (hasCIU OrElse hasTOSI) AndAlso (CIU_PtankContainedInBoots OrElse TOSI_PtankContainedInBoots OrElse CIU_PdispContainedInBoots OrElse TOSI_PdispContainedInBoots) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "3.7.4"
                                    If (hasCIU OrElse hasTOSI) AndAlso (CIU_PtankContainedInSumps OrElse TOSI_PtankContainedInSumps OrElse CIU_PdispContainedInSumps OrElse TOSI_PdispContainedInSumps) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "3.7.5"
                                    If (hasCIU OrElse hasTOSI) AndAlso (CIU_PdispCathodicallyProtected OrElse TOSI_PdispCathodicallyProtected OrElse CIU_PtankCathodicallyProtected OrElse TOSI_PtankCathodicallyProtected) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "3.7.6"
                                    If (hasCIU OrElse hasTOSI) AndAlso (CIU_PdispCathodicallyProtected OrElse TOSI_PdispCathodicallyProtected OrElse CIU_PtankCathodicallyProtected OrElse TOSI_PtankCathodicallyProtected) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                        AddCPReadingsTerm = True
                                        CPReadingsTermQID = oInspCLInfo.ID
                                    End If
                                    'Case "3.7.6.1"
                                    '    If (hasCIU Orelse hasTOSI) Andalso (CIU_PdispCathodicallyProtected Orelse TOSI_PdispCathodicallyProtected Orelse CIU_PtankCathodicallyProtected Orelse TOSI_PtankCathodicallyProtected) Then
                                    '        CPReadingsTermCitationID = oInspCLInfo.Citation
                                    '        CPReadingsTermCitationDesc = oInspCLInfo.DiscrepText
                                    '    End If
                                Case "4"
                                    If hasCIU AndAlso Not CIUemergenOnly Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    ElseIf hasCIU AndAlso CIUgroundWaterVaporMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    ElseIf hasCIU AndAlso CIUinventoryControlPTT Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    ElseIf hasCIU AndAlso CIUautomaticTankGauging Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    ElseIf hasCIU AndAlso CIUstatisticalInventoryReconciliation Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    ElseIf hasCIU AndAlso CIUmanualTankGauging Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    ElseIf hasCIU AndAlso CIUvisualInterstitialMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.1"
                                    If hasCIU AndAlso Not CIUemergenOnly Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.2"
                                    If hasCIU AndAlso CIUgroundWaterVaporMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.2.1"
                                    If hasCIU AndAlso CIUgroundWaterVaporMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.2.2"
                                    If hasCIU AndAlso CIUgroundWaterVaporMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.2.3"
                                    If hasCIU AndAlso CIUgroundWaterVaporMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.2.4"
                                    If hasCIU AndAlso CIUgroundWaterVaporMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.2.5"
                                    If hasCIU AndAlso CIUgroundWaterVaporMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.2.6"
                                    If hasCIU AndAlso CIUgroundWaterVaporMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.2.7"
                                    If hasCIU AndAlso CIUgroundWaterVaporMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.2.8"
                                    If hasCIU AndAlso CIUgroundWaterVaporMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                        If oInspCLInfo.ID > -1 Then
                                            AddMonitorWellsTank = True
                                            MonitorWellsTankQID = oInspCLInfo.ID
                                        End If

                                    End If
                                    'Case "4.2.8.1"
                                    '    If hasCIU Andalso CIUgroundWaterVaporMonitoring Then
                                    '        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    '        AddMonitorWellsTank = True
                                    '        MonitorWellsTankQID = oInspCLInfo.ID
                                    '    End If
                                Case "4.3"
                                    If hasCIU AndAlso CIUinventoryControlPTT Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.3.1"
                                    If hasCIU AndAlso CIUinventoryControlPTT Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.3.2"
                                    If hasCIU AndAlso CIUinventoryControlPTT Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.3.3"
                                    If hasCIU AndAlso CIUinventoryControlPTT Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.3.4"
                                    If hasCIU AndAlso CIUinventoryControlPTT Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.3.5"
                                    If hasCIU AndAlso CIUinventoryControlPTT Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.3.6"
                                    If hasCIU AndAlso CIUinventoryControlPTT Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.4"
                                    If hasCIU AndAlso CIUautomaticTankGauging Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.4.1"
                                    If hasCIU AndAlso CIUautomaticTankGauging Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.4.2"
                                    If hasCIU AndAlso CIUautomaticTankGauging Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.4.3"
                                    If hasCIU AndAlso CIUautomaticTankGauging Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.4.4"
                                    If hasCIU AndAlso CIUautomaticTankGauging Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.4.5"
                                    If hasCIU AndAlso CIUautomaticTankGauging Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.4.5.1"
                                    If hasCIU AndAlso CIUautomaticTankGauging Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If

                                Case "4.5"
                                    If hasCIU AndAlso CIUstatisticalInventoryReconciliation Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.4.6"
                                    If hasCIU AndAlso CIUautomaticTankGauging Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.5.1"
                                    If hasCIU AndAlso CIUstatisticalInventoryReconciliation Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.5.2"
                                    If hasCIU AndAlso CIUstatisticalInventoryReconciliation Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.6"
                                    If hasCIU AndAlso CIUmanualTankGauging Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.6.1"
                                    If hasCIU AndAlso CIUmanualTankGauging Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.6.2"
                                    If hasCIU AndAlso CIUmanualTankGauging Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.6.3"
                                    If hasCIU AndAlso CIUmanualTankGauging Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.6.4"
                                    If hasCIU AndAlso CIUmanualTankGauging Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.7"
                                    If hasCIU AndAlso CIUvisualInterstitialMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.7.1"
                                    If hasCIU AndAlso CIUvisualInterstitialMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.7.2"
                                    If hasCIU AndAlso CIUvisualInterstitialMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.8"
                                    If hasCIU AndAlso CIUelectronicInterstitialMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.8.1"
                                    If hasCIU AndAlso CIUelectronicInterstitialMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.8.2"
                                    If hasCIU AndAlso CIUelectronicInterstitialMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.8.3"
                                    If hasCIU AndAlso CIUelectronicInterstitialMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.8.4"
                                    If hasCIU AndAlso CIUelectronicInterstitialMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "4.8.4.1"
                                    If hasCIU AndAlso CIUelectronicInterstitialMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If

                                Case "4.8.5"
                                    If hasCIU AndAlso CIUelectronicInterstitialMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If

                                Case "5"
                                    If hasCIU AndAlso CIUhasPipes AndAlso Not CIUemergenOnly AndAlso CIU_PpressurizedUSSuction Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    ElseIf hasCIU AndAlso CIU_PgroundWaterVaporMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    ElseIf hasCIU AndAlso CIU_PLTT Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    ElseIf hasCIU AndAlso CIU_PelectronicALLD Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    ElseIf hasCIU AndAlso CIU_PstatisticalInventoryReconciliation Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    ElseIf hasCIU AndAlso (CIU_PvisualInterstitialMonitoring OrElse CIU_PcontinuousInterstitialMonitoring) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    ElseIf hasCIU AndAlso CIU_PvisualInterstitialMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.1"
                                    If hasCIU AndAlso CIUhasPipes AndAlso Not CIUemergenOnly AndAlso CIU_PpressurizedUSSuction Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.2"
                                    If hasCIU AndAlso CIU_PgroundWaterVaporMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.2.1"
                                    If hasCIU AndAlso CIU_PgroundWaterVaporMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.2.2"
                                    If hasCIU AndAlso CIU_PgroundWaterVaporMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.2.3"
                                    If hasCIU AndAlso CIU_PgroundWaterVaporMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.2.4"
                                    If hasCIU AndAlso CIU_PgroundWaterVaporMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.2.5"
                                    If hasCIU AndAlso CIU_PgroundWaterVaporMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.2.6"
                                    If hasCIU AndAlso CIU_PgroundWaterVaporMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.2.7"
                                    If hasCIU AndAlso CIU_PgroundWaterVaporMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.2.8"
                                    If hasCIU AndAlso CIU_PgroundWaterVaporMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                        If oInspCLInfo.ID > -1 Then
                                            AddMonitorWellsPipe = True
                                            MonitorWellsPipeQID = oInspCLInfo.ID
                                        End If
                                    End If
                                    'Case "5.2.8.1"
                                    '    If hasCIU Andalso CIU_PgroundWaterVaporMonitoring Then
                                    '        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    '        AddMonitorWellsPipe = True
                                    '        MonitorWellsPipeQID = oInspCLInfo.ID
                                    '    End If
                                Case "5.3"
                                    If hasCIU AndAlso CIU_PLTT Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.3.1"
                                    If hasCIU AndAlso CIU_PLTT AndAlso CIU_Ppressurized Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.3.2"
                                    If hasCIU AndAlso CIU_PLTT AndAlso CIU_PUSSuction Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.3.3"
                                    If hasCIU AndAlso CIU_PLTT Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.4"
                                    If hasCIU AndAlso CIU_PelectronicALLD Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.4.1"
                                    If hasCIU AndAlso CIU_PelectronicALLD Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.4.2"
                                    If hasCIU AndAlso CIU_PelectronicALLD Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.5"
                                    If hasCIU AndAlso CIU_PstatisticalInventoryReconciliation Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.5.1"
                                    If hasCIU AndAlso CIU_PstatisticalInventoryReconciliation Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.5.2"
                                    If hasCIU AndAlso CIU_PstatisticalInventoryReconciliation Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.6"
                                    If hasCIU AndAlso (CIU_PvisualInterstitialMonitoring OrElse CIU_PcontinuousInterstitialMonitoring) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.6.1"
                                    If hasCIU AndAlso (CIU_PvisualInterstitialMonitoring OrElse CIU_PcontinuousInterstitialMonitoring) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.6.2"
                                    If hasCIU AndAlso (CIU_PvisualInterstitialMonitoring OrElse CIU_PcontinuousInterstitialMonitoring) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.6.3"
                                    If hasCIU AndAlso (CIU_PvisualInterstitialMonitoring OrElse CIU_PcontinuousInterstitialMonitoring) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.6.4"
                                    If hasCIU AndAlso (CIU_PvisualInterstitialMonitoring OrElse CIU_PcontinuousInterstitialMonitoring) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.7"
                                    If hasCIU AndAlso CIU_PvisualInterstitialMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.7.1"
                                    If hasCIU AndAlso CIU_PvisualInterstitialMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.7.2"
                                    If hasCIU AndAlso CIU_PvisualInterstitialMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.7.2.1"
                                    If hasCIU AndAlso CIU_PvisualInterstitialMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If

                                Case "5.8"
                                    If hasCIU AndAlso CIU_PcontinuousInterstitialMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.8.1"
                                    If hasCIU AndAlso CIU_PcontinuousInterstitialMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.8.2"
                                    If hasCIU AndAlso CIU_PcontinuousInterstitialMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.8.3"
                                    If hasCIU AndAlso CIU_PcontinuousInterstitialMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.8.4"
                                    If hasCIU AndAlso CIU_PcontinuousInterstitialMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.8.4.1"
                                    If hasCIU AndAlso CIU_PcontinuousInterstitialMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If

                                Case "5.8.5"
                                    If hasCIU AndAlso CIU_PcontinuousInterstitialMonitoring Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If

                                Case "5.9"
                                    If hasCIU AndAlso CIU_Ppressurized AndAlso Not CIUemergenOnly Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    ElseIf hasCIU AndAlso CIU_Pelectronic Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.9.1"
                                    If hasCIU AndAlso CIU_Ppressurized AndAlso Not CIUemergenOnly Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.9.2"
                                    If hasCIU AndAlso CIU_Ppressurized AndAlso Not CIUemergenOnly Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.9.3"
                                    If hasCIU AndAlso CIU_Ppressurized AndAlso Not CIUemergenOnly Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.9.4"
                                    If hasCIU AndAlso CIU_Ppressurized AndAlso Not CIUemergenOnly AndAlso (CIU_Pmechanical Or Me.CIU_Pelectronic) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.9.5"
                                    If hasCIU AndAlso CIU_Ppressurized AndAlso Not CIUemergenOnly AndAlso CIU_Pmechanical Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "5.9.6"
                                    If hasCIU AndAlso CIU_Pelectronic Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "6"
                                    If hasCIU AndAlso Not CIUUsedOilOnly Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    ElseIf hasCIU AndAlso Not (CIUUsedOilOnly OrElse CIUemergenOnly) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    ElseIf hasCIU AndAlso CIU_PdispContainedInSumps Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    ElseIf hasCIU AndAlso CIU_Ppressurized Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    ElseIf hasCIU AndAlso CIU_PtankContainedInSumps Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    ElseIf hasCIU AndAlso CIU_Pplastic Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "6.1"
                                    If hasCIU AndAlso Not (CIUUsedOilOnly OrElse CIUemergenOnly) Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "6.2"
                                    If hasCIU AndAlso CIU_PdispContainedInSumps Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "6.2.1"
                                    If hasCIU AndAlso CIU_PdispContainedInSumps AndAlso CIU_PipeInstalledAfter10_1_08 Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If

                                Case "6.3"
                                    If hasCIU AndAlso CIU_Ppressurized Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "6.4"
                                    If hasCIU AndAlso CIU_PtankContainedInSumps Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "6.4.1"
                                    If hasCIU AndAlso CIU_PtankContainedInSumps AndAlso CIU_PipeInstalledAfter10_1_08 Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If

                                Case "6.5"
                                    If hasCIU AndAlso CIU_Pplastic Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "7"
                                    If hasTOS OrElse hasTOSI Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "7.1"
                                    If hasTOS OrElse hasTOSI Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "7.2"
                                    If hasTOS OrElse hasTOSI Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "7.3"
                                    If hasTOS OrElse hasTOSI Then
                                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    End If
                                Case "8"
                                    ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                Case "9"
                                    ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                Case "11"
                                    ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])
                                    MonitorWellsTankPipeQID = oInspCLInfo.ID
                            End Select
                        End If
                        '    End If
                    Next
                Else

                    If Not oInspection.ChecklistMasterCollection Is Nothing AndAlso oInspection.ChecklistMasterCollection.Count > 0 Then
                        oInspection.ChecklistMasterCollection.Clear()
                        RetrieveAllChecklistText(oInspection.FacilityID)
                    End If

                    For Each oInspCLInfo In oInspection.ChecklistMasterCollection.Values




                        oInspCLInfo.Show = False

                        ShowCLItem(oInspCLInfo, True, hasResponses, [readOnly])

                        Select Case (oInspCLInfo.CheckListItemNumber.Trim)

                            Case "2.6"
                                qIDOfCLItem2point6 = oInspCLInfo.ID
                            Case "3.5.4"

                                AddCPReadingsTank = True
                                CPReadingsTankQID = oInspCLInfo.ID
                            Case "3.6.3"
                                AddCPReadingsPipe = True
                                CPReadingsPipeQID = oInspCLInfo.ID
                            Case "3.7.6"
                                AddCPReadingsTerm = True
                                CPReadingsTermQID = oInspCLInfo.ID
                            Case "4.2.8"

                                If oInspCLInfo.ID > -1 Then
                                    AddMonitorWellsTank = True
                                    MonitorWellsTankQID = oInspCLInfo.ID
                                End If

                            Case "5.2.8"
                                If oInspCLInfo.ID > -1 Then

                                    AddMonitorWellsPipe = True
                                    MonitorWellsPipeQID = oInspCLInfo.ID

                                End If
                            Case "11"
                                MonitorWellsTankPipeQID = oInspCLInfo.ID
                        End Select

                    Next

                End If


                CheckTankPipeBelongToCL(, , [readOnly])
                AddCPReadings(hasResponses, [readOnly])
                AddMonitorWells(hasResponses, [readOnly])
                AddSOC(hasResponses, [readOnly])
                'CheckTankPipeBelongToCL(, , [readOnly])
                If Not Me.oInspectionChecklistMasterDB.GetDBInspctionCheckListApproval Then
                    AddCCAT(hasResponses, [readOnly])
                Else
                    AddCCAT2(hasResponses, [readOnly])
                End If
                ' citation / discrep is required only if there are any responses with answer no
                ' when you first generate the checklist, there are no responses with answer no
                ' hence no need to call the below two functions when checklist is generated
                AddCitation(hasResponses, [readOnly])
                AddDiscrep(hasResponses, [readOnly])
                AddComments(hasResponses, [readOnly])
                AddTankFuelType()
                AddPipeFuelType()
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub ShowCLItem(ByRef oInspCLInfo As MUSTER.Info.InspectionChecklistMasterInfo, ByVal bolShow As Boolean, ByVal hasResponses As Boolean, Optional ByVal [readOnly] As Boolean = False)
            Try
                Dim oInspRespInfo As MUSTER.Info.InspectionResponsesInfo
                If Not hasResponses And Not [readOnly] And bolShow Then
                    oInspRespInfo = New MUSTER.Info.InspectionResponsesInfo(0, _
                                        oInspection.ID, _
                                        oInspCLInfo.ID, _
                                        IIf(oInspCLInfo.SOC = String.Empty, False, True), _
                                        -1, _
                                        False, _
                                        String.Empty, _
                                        CDate("01/01/0001"), _
                                        String.Empty, _
                                        CDate("01/01/0001"))
                    oInspectionResponses.Add(oInspRespInfo)
                Else
                    Dim bolResponseFound As Boolean = False
                    For Each respInfo As MUSTER.Info.InspectionResponsesInfo In oInspection.ResponsesCollection.Values
                        If respInfo.QuestionID = oInspCLInfo.ID And respInfo.InspectionID = oInspection.ID Then
                            bolResponseFound = True
                            respInfo.Deleted = Not bolShow
                            If oInspCLInfo.CheckListItemNumber = "2.5" Then
                                respOfCLItem2point5 = respInfo.Response
                            ElseIf oInspCLInfo.CheckListItemNumber = "2.6" And [readOnly] And respInfo.Response = -1 Then
                                respInfo.Response = -2
                            ElseIf oInspCLInfo.CheckListItemNumber = "2.6" And Not [readOnly] And respInfo.Response = -2 Then
                                respInfo.Response = -1

                            End If
                        End If
                    Next
                    If Not bolResponseFound And (Not [readOnly] Or oInspCLInfo.CheckListItemNumber.StartsWith("12.") Or oInspCLInfo.CheckListItemNumber = "12") And bolShow Then
                        oInspRespInfo = New MUSTER.Info.InspectionResponsesInfo(0, _
                                            oInspection.ID, _
                                            oInspCLInfo.ID, _
                                            IIf(oInspCLInfo.SOC = String.Empty, False, True), _
                                            IIf(Not [readOnly], -1, 2), _
                                            False, _
                                            String.Empty, _
                                            CDate("01/01/0001"), _
                                            String.Empty, _
                                            CDate("01/01/0001"))
                        oInspectionResponses.Add(oInspRespInfo)
                        If oInspCLInfo.CheckListItemNumber = "2.5" Then
                            respOfCLItem2point5 = oInspRespInfo.Response
                        ElseIf oInspCLInfo.CheckListItemNumber = "2.6" And [readOnly] And oInspRespInfo.Response = -1 Then
                            oInspRespInfo.Response = -2
                        ElseIf oInspCLInfo.CheckListItemNumber = "2.6" And Not [readOnly] And oInspRespInfo.Response = -2 Then
                            oInspRespInfo.Response = -1

                        End If
                    End If
                End If
                If oInspCLInfo.ID > -1 Then
                    oInspCLInfo.Show = bolShow
                End If

            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Private Sub AddRectifier(ByVal hasResponses As Boolean, ByVal qid As Int64, Optional ByVal [readOnly] As Boolean = False)
            Dim bolAddRec As Boolean = True
            Try
                If hasResponses Then
                    If oInspection.RectifiersCollection.Count > 0 Then
                        For Each rectifier As MUSTER.Info.InspectionRectifierInfo In oInspection.RectifiersCollection.Values
                            If rectifier.QuestionID = qid And rectifier.InspectionID = oInspection.ID And rectifier.Deleted Then
                                bolAddRec = False
                                rectifier.Deleted = False
                            End If
                        Next
                    End If
                End If
                If bolAddRec And Not [readOnly] Then
                    Dim rectifier As New MUSTER.Info.InspectionRectifierInfo(0, _
                                            oInspection.ID, _
                                            qid, _
                                            False, _
                                            String.Empty, _
                                            0.0, _
                                            0.0, _
                                            0.0, _
                                            False, _
                                            String.Empty, _
                                            CDate("01/01/0001"), _
                                            String.Empty, _
                                            CDate("01/01/0001"))
                    oInspectionRectifier.InspectionInfo = oInspection
                    oInspectionRectifier.Add(rectifier)
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub AddCPReadings(ByVal hasResponses As Boolean, Optional ByVal [readOnly] As Boolean = False)
            Try
                'Dim tnk As MUSTER.Info.TankInfo
                'Dim pipe As MUSTER.Info.PipeInfo
                'Dim bolAddTankCP, bolAddPipeCP, bolAddTermCP As Boolean
                Dim cpreading As MUSTER.Info.InspectionCPReadingsInfo
                Dim clInfo As MUSTER.Info.InspectionChecklistMasterInfo
                'Dim alTank As New ArrayList
                Dim slTank As New SortedList
                Dim alPipe1 As New ArrayList
                'Dim alPipe2 As New ArrayList
                Dim slPipe1 As New SortedList
                'Dim slPipe2 As New SortedList
                Dim alTerm1 As New ArrayList
                'Dim alTerm2 As New ArrayList
                'Dim alTerm3 As New ArrayList
                Dim slTerm1 As New SortedList
                'Dim slTerm2 As New SortedList
                'Dim slTerm3 As New SortedList
                Dim slTankWithResponseNo As New SortedList
                Dim slPipeWithResponseNo As New SortedList
                Dim slTermWithResponseNo As New SortedList
                Dim bolAddTankRemote As Boolean = True
                Dim bolAddPipeRemote As Boolean = True
                Dim bolAddTermRemote As Boolean = True
                Dim bolAddTankGalvanic As Boolean = True
                Dim bolAddPipeGalvanic As Boolean = True
                Dim bolAddTermGalvanic As Boolean = True
                Dim bolAddTankInspectorTested As Boolean = True
                Dim bolAddPipeInspectorTested As Boolean = True
                Dim bolAddTermInspectorTested As Boolean = True
                Dim nTankInspectorTestedID, nPipeInspectorTestedID, nTermInspectorTestedID As Integer
                Dim bolTankInspectorTestedResponse As Boolean = True
                Dim bolPipeInspectorTestedResponse As Boolean = True
                Dim bolTermInspectorTestedResponse As Boolean = True
                oInspectionCPReadings.InspectionInfo = oInspection

                If AddCPReadingsTank And CPReadingsTankQID <> 0 Then
                    clInfo = Nothing
                    clInfo = oInspection.ChecklistMasterCollection.Item(CPReadingsTankQID)
                    If Not clInfo Is Nothing Then
                        If clInfo.Show Then
                            slTankID = New SortedList
                            For Each nTank As Integer In clInfo.TankArrayList.ToArray
                                slTankID.Add(nTank, htTankIDIndex.Item(nTank))
                                'alTank.Add(nTank)
                                slTank.Add(htTankIDIndex.Item(nTank), nTank)
                            Next
                            If hasResponses Then
                                For Each cp As MUSTER.Info.InspectionCPReadingsInfo In oInspection.CPReadingsCollection.Values
                                    If cp.QuestionID = CPReadingsTankQID Then
                                        If cp.RemoteReferCellPlacement Then
                                            bolAddTankRemote = False
                                            cp.Deleted = False
                                        ElseIf cp.GalvanicIC Then
                                            bolAddTankGalvanic = False
                                            cp.Deleted = False
                                        ElseIf cp.TestedByInspector Then
                                            nTankInspectorTestedID = cp.ID
                                            bolAddTankInspectorTested = False
                                            cp.Deleted = False
                                        Else
                                            If slTank.Contains(cp.TankPipeIndex) And cp.TankPipeEntityID = 12 Then
                                                If cp.Deleted Then
                                                    slTank.Remove(cp.TankPipeIndex)
                                                Else
                                                    'cp.Deleted = False
                                                    slTank.Remove(cp.TankPipeIndex)
                                                    If cp.PassFailIncon = 0 Or cp.PassFailIncon = 2 Then
                                                        If Not slTankWithResponseNo.Contains(cp.TankPipeIndex) Then
                                                            slTankWithResponseNo.Add(cp.TankPipeIndex, cp.TankPipeIndex)
                                                        End If
                                                    End If
                                                End If
                                            ElseIf clInfo.TankArrayList.Contains(cp.TankPipeID) And cp.TankPipeEntityID = 12 Then
                                                If Not cp.Deleted Then
                                                    'cp.Deleted = False
                                                    If Not slTankID.Contains(cp.TankPipeID) Then
                                                        slTankID.Add(cp.TankPipeID, htTankIDIndex.Item(cp.TankPipeID))
                                                    End If
                                                    If cp.PassFailIncon = 0 Or cp.PassFailIncon = 2 Then
                                                        If Not slTankWithResponseNo.Contains(cp.TankPipeIndex) Then
                                                            slTankWithResponseNo.Add(cp.TankPipeIndex, cp.TankPipeIndex)
                                                        End If
                                                    End If
                                                End If
                                            ElseIf Not [readOnly] Then
                                                cp.Deleted = True
                                            End If

                                            'If alTank.Contains(cp.TankPipeID) And cp.TankPipeEntityID = 12 Then
                                            '    cp.Deleted = False
                                            '    alTank.Remove(cp.TankPipeID)
                                            '    If cp.PassFailIncon = 0 Or cp.PassFailIncon = 2 Then
                                            '        If Not slTankWithResponseNo.Contains(cp.TankPipeIndex) Then
                                            '            slTankWithResponseNo.Add(cp.TankPipeIndex, cp.TankPipeIndex)
                                            '        End If
                                            '    End If
                                            'ElseIf clInfo.TankArrayList.Contains(cp.TankPipeID) And cp.TankPipeEntityID = 12 Then
                                            '    cp.Deleted = False
                                            '    If cp.PassFailIncon = 0 Or cp.PassFailIncon = 2 Then
                                            '        If Not slTankWithResponseNo.Contains(cp.TankPipeIndex) Then
                                            '            slTankWithResponseNo.Add(cp.TankPipeIndex, cp.TankPipeIndex)
                                            '        End If
                                            '    End If
                                            'End If
                                        End If
                                    End If
                                Next
                            End If ' If hasResponses Then

                            If bolAddTankInspectorTested And Not [readOnly] Then
                                cpreading = New MUSTER.Info.InspectionCPReadingsInfo(0, _
                                            oInspection.ID, _
                                            CPReadingsTankQID, _
                                            0, _
                                            0, _
                                            0, _
                                            String.Empty, _
                                            String.Empty, _
                                            String.Empty, _
                                            String.Empty, _
                                            -1, _
                                            False, _
                                            String.Empty, _
                                            CDate("01/01/0001"), _
                                            String.Empty, _
                                            CDate("01/01/0001"), _
                                            -1, _
                                            False, _
                                            False, _
                                            -1, True, True)
                                oInspectionCPReadings.Add(cpreading)
                            ElseIf Not bolAddTankInspectorTested And Not [readOnly] Then
                                If nTankInspectorTestedID <> 0 Then
                                    If Not oInspection.CPReadingsCollection.Item(nTankInspectorTestedID) Is Nothing Then
                                        If oInspection.CPReadingsCollection.Item(nTankInspectorTestedID).TestedByInspectorResponse = False Then
                                            bolTankInspectorTestedResponse = False
                                        End If
                                    End If
                                End If
                            End If

                            If bolAddTankRemote And Not [readOnly] Then
                                cpreading = New MUSTER.Info.InspectionCPReadingsInfo(0, _
                                            oInspection.ID, _
                                            CPReadingsTankQID, _
                                            0, _
                                            0, _
                                            0, _
                                            String.Empty, _
                                            String.Empty, _
                                            String.Empty, _
                                            String.Empty, _
                                            -1, _
                                            False, _
                                            String.Empty, _
                                            CDate("01/01/0001"), _
                                            String.Empty, _
                                            CDate("01/01/0001"), _
                                            -1, _
                                            True, _
                                            False, _
                                            -1, False, False)
                                oInspectionCPReadings.Add(cpreading)
                            End If

                            If bolAddTankGalvanic And Not [readOnly] Then
                                cpreading = New MUSTER.Info.InspectionCPReadingsInfo(0, _
                                            oInspection.ID, _
                                            CPReadingsTankQID, _
                                            0, _
                                            0, _
                                            0, _
                                            String.Empty, _
                                            String.Empty, _
                                            String.Empty, _
                                            String.Empty, _
                                            -1, _
                                            False, _
                                            String.Empty, _
                                            CDate("01/01/0001"), _
                                            String.Empty, _
                                            CDate("01/01/0001"), _
                                            0, _
                                            False, _
                                            True, _
                                            -1, False, False)
                                oInspectionCPReadings.Add(cpreading)
                            End If

                            If slTank.Count > 0 And Not [readOnly] Then
                                For i As Integer = 0 To slTank.Count - 1
                                    cpreading = New MUSTER.Info.InspectionCPReadingsInfo(0, _
                                                oInspection.ID, _
                                                CPReadingsTankQID, _
                                                slTank.GetByIndex(i), _
                                                slTank.GetKey(i), _
                                                12, _
                                                String.Empty, _
                                                String.Empty, _
                                                String.Empty, _
                                                String.Empty, _
                                                -1, _
                                                False, _
                                                String.Empty, _
                                                CDate("01/01/0001"), _
                                                String.Empty, _
                                                CDate("01/01/0001"), _
                                                nMaxTankCPLineNum, _
                                                False, _
                                                False, _
                                                -1, False, False)
                                    oInspectionCPReadings.Add(cpreading)
                                    nMaxTankCPLineNum += 1
                                Next
                            End If ' If slTank.Count > 0 And Not [readOnly] Then
                            'If alTank.Count > 0 And Not [readOnly] Then
                            '    For Each nTank As Integer In alTank.ToArray
                            '        cpreading = New MUSTER.Info.InspectionCPReadingsInfo(0, _
                            '                    oInspection.ID, _
                            '                    CPReadingsTankQID, _
                            '                    nTank, _
                            '                    slTankID.Item(ntank), _
                            '                    12, _
                            '                    String.Empty, _
                            '                    String.Empty, _
                            '                    String.Empty, _
                            '                    String.Empty, _
                            '                    -1, _
                            '                    False, _
                            '                    String.Empty, _
                            '                    CDate("01/01/0001"), _
                            '                    String.Empty, _
                            '                    CDate("01/01/0001"), _
                            '                    nMaxTankCPLineNum, _
                            '                    False, _
                            '                    False, _
                            '                    -1)
                            '        oInspectionCPReadings.Add(cpreading)
                            '        nMaxTankCPLineNum += 1
                            '    Next
                            'End If ' If alTank.Count > 0 And Not [readOnly] Then

                            ' if there were any cp readings with response no, check if citation exists and mark deleted
                            ' else create citation
                            If slTankWithResponseNo.Count > 0 Then
                                Dim citation As MUSTER.Info.InspectionCitationInfo
                                If bolTankInspectorTestedResponse Then
                                    Dim createCitation As Boolean = True
                                    Dim strTanks As String = String.Empty

                                    For i As Integer = 0 To slTankWithResponseNo.Count - 1
                                        strTanks += "T" + slTankWithResponseNo.GetByIndex(i).ToString + ", "
                                    Next
                                    If strTanks <> String.Empty Then
                                        strTanks = strTanks.Trim.TrimEnd(",")
                                    End If

                                    For Each Citation In oInspection.CitationsCollection.Values
                                        If Citation.QuestionID = CPReadingsTankQID Then
                                            Citation.Deleted = False
                                            createCitation = False
                                            Citation.CCAT = strTanks
                                            Exit For
                                        End If
                                    Next
                                    If createCitation Then
                                        citation = New MUSTER.Info.InspectionCitationInfo(0, _
                                        oInspection.ID, _
                                        CPReadingsTankQID, _
                                        oInspection.FacilityID, _
                                        0, _
                                        0, _
                                        clInfo.Citation, _
                                        strTanks, _
                                        False, _
                                        CDate("01/01/0001"), _
                                        CDate("01/01/0001"), _
                                        CDate("01/01/0001"), _
                                        False, _
                                        String.Empty, _
                                        CDate("01/01/0001"), _
                                        String.Empty, _
                                        CDate("01/01/0001"))
                                        oInspectionCitation.InspectionInfo = oInspection
                                        oInspectionCitation.Add(citation)
                                    End If ' If createCitation Then
                                Else
                                    ' mark citations deleted
                                    For Each Citation In oInspection.CitationsCollection.Values
                                        If Citation.QuestionID = CPReadingsTankQID Then
                                            Citation.Deleted = True
                                            Exit For
                                        End If
                                    Next
                                    For Each discrep As MUSTER.Info.InspectionDiscrepInfo In oInspection.DiscrepsCollection.Values
                                        If discrep.QuestionID = CPReadingsTankQID Then
                                            discrep.Deleted = True
                                            Exit For
                                        End If
                                    Next
                                End If
                            End If ' If slTankWithResponseNo.Count > 0 Then

                        End If ' If clInfo.Show Then
                    End If ' If Not clInfo Is Nothing Then
                End If

                If AddCPReadingsPipe And CPReadingsPipeQID <> 0 Then
                    clInfo = Nothing
                    clInfo = oInspection.ChecklistMasterCollection.Item(CPReadingsPipeQID)
                    If Not clInfo Is Nothing Then
                        If clInfo.Show Then
                            slPipeID = New SortedList
                            For Each nPipe As Integer In clInfo.PipeArrayList.ToArray
                                slPipeID.Add(npipe, htPipeIDIndex.Item(npipe))
                                alPipe1.Add(nPipe)
                                'alPipe2.Add(npipe)
                                slPipe1.Add(htPipeIDIndex.Item(npipe), npipe)
                                'slPipe2.Add(htPipeIDIndex.Item(npipe), npipe)
                            Next
                            If hasResponses Then
                                For Each cp As MUSTER.Info.InspectionCPReadingsInfo In oInspection.CPReadingsCollection.Values
                                    If cp.QuestionID = CPReadingsPipeQID Then
                                        If cp.RemoteReferCellPlacement Then
                                            cp.Deleted = False
                                            bolAddPipeRemote = False
                                        ElseIf cp.GalvanicIC Then
                                            bolAddPipeGalvanic = False
                                            cp.Deleted = False
                                        ElseIf cp.TestedByInspector Then
                                            nPipeInspectorTestedID = cp.ID
                                            bolAddPipeInspectorTested = False
                                            cp.Deleted = False
                                        Else
                                            If slPipe1.Contains(cp.TankPipeIndex) And cp.TankPipeEntityID = 10 Then
                                                If cp.Deleted Then
                                                    slPipe1.Remove(cp.TankPipeIndex)
                                                Else
                                                    'cp.Deleted = False
                                                    slPipe1.Remove(cp.TankPipeIndex)
                                                    If cp.PassFailIncon = 0 Or cp.PassFailIncon = 2 Then
                                                        If Not slPipeWithResponseNo.Contains(cp.TankPipeIndex) Then
                                                            slPipeWithResponseNo.Add(cp.TankPipeIndex, cp.TankPipeIndex)
                                                        End If
                                                    End If
                                                End If
                                                'ElseIf slPipe2.Contains(cp.TankPipeIndex) And cp.TankPipeEntityID = 10 Then
                                                '    cp.Deleted = False
                                                '    slPipe2.Remove(cp.TankPipeIndex)
                                                '    If cp.PassFailIncon = 0 Or cp.PassFailIncon = 2 Then
                                                '        If Not slPipeWithResponseNo.Contains(cp.TankPipeIndex) Then
                                                '            slPipeWithResponseNo.Add(cp.TankPipeIndex, cp.TankPipeIndex)
                                                '        End If
                                                '    End If
                                            ElseIf clInfo.PipeArrayList.Contains(cp.TankPipeID) And cp.TankPipeEntityID = 10 Then
                                                If Not cp.Deleted Then
                                                    'cp.Deleted = False
                                                    If Not slPipeID.Contains(cp.TankPipeID) Then
                                                        slPipeID.Add(cp.TankPipeID, htPipeIDIndex.Item(cp.TankPipeID))
                                                    End If
                                                    If cp.PassFailIncon = 0 Or cp.PassFailIncon = 2 Then
                                                        If Not slPipeWithResponseNo.Contains(cp.TankPipeIndex) Then
                                                            slPipeWithResponseNo.Add(cp.TankPipeIndex, cp.TankPipeIndex)
                                                        End If
                                                    End If
                                                End If
                                            ElseIf Not [readOnly] Then
                                                cp.Deleted = True
                                            End If
                                            'If alPipe1.Contains(cp.TankPipeID) And cp.TankPipeEntityID = 10 Then
                                            '    cp.Deleted = False
                                            '    alPipe1.Remove(cp.TankPipeID)
                                            '    If cp.PassFailIncon = 0 Or cp.PassFailIncon = 2 Then
                                            '        If Not slPipeWithResponseNo.Contains(cp.TankPipeIndex) Then
                                            '            slPipeWithResponseNo.Add(cp.TankPipeIndex, cp.TankPipeIndex)
                                            '        End If
                                            '    End If
                                            'ElseIf alPipe2.Contains(cp.TankPipeID) And cp.TankPipeEntityID = 10 Then
                                            '    cp.Deleted = False
                                            '    alPipe2.Remove(cp.TankPipeID)
                                            '    If cp.PassFailIncon = 0 Or cp.PassFailIncon = 2 Then
                                            '        If Not slPipeWithResponseNo.Contains(cp.TankPipeIndex) Then
                                            '            slPipeWithResponseNo.Add(cp.TankPipeIndex, cp.TankPipeIndex)
                                            '        End If
                                            '    End If
                                            'ElseIf clInfo.PipeArrayList.Contains(cp.TankPipeID) And cp.TankPipeEntityID = 10 Then
                                            '    cp.Deleted = False
                                            '    If cp.PassFailIncon = 0 Or cp.PassFailIncon = 2 Then
                                            '        If Not slPipeWithResponseNo.Contains(cp.TankPipeIndex) Then
                                            '            slPipeWithResponseNo.Add(cp.TankPipeIndex, cp.TankPipeIndex)
                                            '        End If
                                            '    End If
                                            'End If
                                        End If
                                    End If
                                Next
                            End If ' If hasResponses Then

                            If bolAddPipeInspectorTested And Not [readOnly] Then
                                cpreading = New MUSTER.Info.InspectionCPReadingsInfo(0, _
                                            oInspection.ID, _
                                            CPReadingsPipeQID, _
                                            0, _
                                            0, _
                                            0, _
                                            String.Empty, _
                                            String.Empty, _
                                            String.Empty, _
                                            String.Empty, _
                                            -1, _
                                            False, _
                                            String.Empty, _
                                            CDate("01/01/0001"), _
                                            String.Empty, _
                                            CDate("01/01/0001"), _
                                            -1, _
                                            False, _
                                            False, _
                                            -1, True, True)
                                oInspectionCPReadings.Add(cpreading)
                            ElseIf Not bolAddPipeInspectorTested And Not [readOnly] Then
                                If nPipeInspectorTestedID <> 0 Then
                                    If Not oInspection.CPReadingsCollection.Item(nPipeInspectorTestedID) Is Nothing Then
                                        If oInspection.CPReadingsCollection.Item(nPipeInspectorTestedID).TestedByInspectorResponse = False Then
                                            bolPipeInspectorTestedResponse = False
                                        End If
                                    End If
                                End If
                            End If

                            If bolAddPipeRemote And Not [readOnly] Then
                                cpreading = New MUSTER.Info.InspectionCPReadingsInfo(0, _
                                            oInspection.ID, _
                                            CPReadingsPipeQID, _
                                            0, _
                                            0, _
                                            0, _
                                            String.Empty, _
                                            String.Empty, _
                                            String.Empty, _
                                            String.Empty, _
                                            -1, _
                                            False, _
                                            String.Empty, _
                                            CDate("01/01/0001"), _
                                            String.Empty, _
                                            CDate("01/01/0001"), _
                                            -1, _
                                            True, _
                                            False, _
                                            -1, False, False)
                                oInspectionCPReadings.Add(cpreading)
                            End If

                            If bolAddPipeGalvanic And Not [readOnly] Then
                                cpreading = New MUSTER.Info.InspectionCPReadingsInfo(0, _
                                            oInspection.ID, _
                                            CPReadingsPipeQID, _
                                            0, _
                                            0, _
                                            0, _
                                            String.Empty, _
                                            String.Empty, _
                                            String.Empty, _
                                            String.Empty, _
                                            -1, _
                                            False, _
                                            String.Empty, _
                                            CDate("01/01/0001"), _
                                            String.Empty, _
                                            CDate("01/01/0001"), _
                                            0, _
                                            False, _
                                            True, _
                                            -1, False, False)
                                oInspectionCPReadings.Add(cpreading)
                            End If

                            If slPipe1.Count > 0 And Not [readOnly] Then
                                'If (slPipe1.Count > 0 Or slPipe2.Count > 0) And Not [readOnly] Then
                                Dim slPipe As New SortedList
                                Dim addPipeNos As Integer = 0
                                For i As Integer = 0 To slPipe1.Count - 1
                                    slPipe.Add(slPipe1.GetKey(i), 1)
                                Next
                                'For i As Integer = 0 To slPipe2.Count - 1
                                '    If slPipe.Contains(slPipe2.GetKey(i)) Then
                                '        slPipe.Item(slPipe2.GetKey(i)) += 1
                                '    Else
                                '        slPipe.Add(slPipe2.GetKey(i), 1)
                                '    End If
                                'Next
                                For i As Integer = 0 To slPipe.Count - 1
                                    addPipeNos = slPipe.GetByIndex(i)
                                    For j As Integer = 0 To addPipeNos - 1
                                        cpreading = New MUSTER.Info.InspectionCPReadingsInfo(0, _
                                                    oInspection.ID, _
                                                    CPReadingsPipeQID, _
                                                    slPipe1.Item(slPipe.GetKey(i)), _
                                                    slPipe.GetKey(i), _
                                                    10, _
                                                    String.Empty, _
                                                    String.Empty, _
                                                    String.Empty, _
                                                    String.Empty, _
                                                    -1, _
                                                    False, _
                                                    String.Empty, _
                                                    CDate("01/01/0001"), _
                                                    String.Empty, _
                                                    CDate("01/01/0001"), _
                                                    nMaxPipeCPLineNum, _
                                                    False, _
                                                    False, _
                                                    -1, False, False)
                                        oInspectionCPReadings.Add(cpreading)
                                        nMaxPipeCPLineNum += 1
                                    Next
                                Next
                            End If ' If (slPipe1.Count > 0 Or slPipe2.Count > 0) And Not [readOnly] Then
                            'If alPipe1.Count > 0 And Not [readOnly] Then
                            '    For Each nPipe As Integer In alPipe1.ToArray
                            '        cpreading = New MUSTER.Info.InspectionCPReadingsInfo(0, _
                            '                    oInspection.ID, _
                            '                    CPReadingsPipeQID, _
                            '                    nPipe, _
                            '                    slPipeID.Item(nPipe), _
                            '                    10, _
                            '                    String.Empty, _
                            '                    String.Empty, _
                            '                    String.Empty, _
                            '                    String.Empty, _
                            '                    -1, _
                            '                    False, _
                            '                    String.Empty, _
                            '                    CDate("01/01/0001"), _
                            '                    String.Empty, _
                            '                    CDate("01/01/0001"), _
                            '                    nMaxPipeCPLineNum, _
                            '                    False, _
                            '                    False, _
                            '                    -1)
                            '        oInspectionCPReadings.Add(cpreading)
                            '        nMaxPipeCPLineNum += 1
                            '    Next
                            'End If ' If alPipe1.Count > 0 And Not [readOnly] Then
                            'If alPipe2.Count > 0 And Not [readOnly] Then
                            '    For Each nPipe As Integer In alPipe1.ToArray
                            '        cpreading = New MUSTER.Info.InspectionCPReadingsInfo(0, _
                            '                    oInspection.ID, _
                            '                    CPReadingsPipeQID, _
                            '                    nPipe, _
                            '                    slPipeID.Item(nPipe), _
                            '                    10, _
                            '                    String.Empty, _
                            '                    String.Empty, _
                            '                    String.Empty, _
                            '                    String.Empty, _
                            '                    -1, _
                            '                    False, _
                            '                    String.Empty, _
                            '                    CDate("01/01/0001"), _
                            '                    String.Empty, _
                            '                    CDate("01/01/0001"), _
                            '                    nMaxPipeCPLineNum, _
                            '                    False, _
                            '                    False, _
                            '                    -1)
                            '        oInspectionCPReadings.Add(cpreading)
                            '        nMaxPipeCPLineNum += 1
                            '    Next
                            'End If ' If alPipe2.Count > 0 And Not [readOnly] Then

                            ' if there were any cp readings with response no, check if citation exists and mark deleted
                            ' else create citation
                            If slPipeWithResponseNo.Count > 0 Then
                                Dim citation As MUSTER.Info.InspectionCitationInfo
                                If bolPipeInspectorTestedResponse Then
                                    Dim createCitation As Boolean = True
                                    Dim strPipes As String = String.Empty

                                    For i As Integer = 0 To slPipeWithResponseNo.Count - 1
                                        strPipes += "P" + slPipeWithResponseNo.GetByIndex(i).ToString + ", "
                                    Next
                                    If strPipes <> String.Empty Then
                                        strPipes = strPipes.Trim.TrimEnd(",")
                                    End If

                                    For Each Citation In oInspection.CitationsCollection.Values
                                        If Citation.QuestionID = CPReadingsPipeQID Then
                                            Citation.Deleted = False
                                            createCitation = False
                                            Citation.CCAT = strPipes
                                            Exit For
                                        End If
                                    Next
                                    If createCitation Then
                                        citation = New MUSTER.Info.InspectionCitationInfo(0, _
                                        oInspection.ID, _
                                        CPReadingsPipeQID, _
                                        oInspection.FacilityID, _
                                        0, _
                                        0, _
                                        clInfo.Citation, _
                                        strPipes, _
                                        False, _
                                        CDate("01/01/0001"), _
                                        CDate("01/01/0001"), _
                                        CDate("01/01/0001"), _
                                        False, _
                                        String.Empty, _
                                        CDate("01/01/0001"), _
                                        String.Empty, _
                                        CDate("01/01/0001"))
                                        oInspectionCitation.InspectionInfo = oInspection
                                        oInspectionCitation.Add(citation)
                                    End If ' If createCitation Then
                                Else
                                    ' mark citations deleted
                                    For Each Citation In oInspection.CitationsCollection.Values
                                        If Citation.QuestionID = CPReadingsPipeQID Then
                                            Citation.Deleted = True
                                            Exit For
                                        End If
                                    Next
                                    For Each discrep As MUSTER.Info.InspectionDiscrepInfo In oInspection.DiscrepsCollection.Values
                                        If discrep.QuestionID = CPReadingsPipeQID Then
                                            discrep.Deleted = True
                                            Exit For
                                        End If
                                    Next
                                End If
                            End If ' If slPipeWithResponseNo.Count > 0 Then

                        End If ' If clInfo.Show Then
                    End If ' If Not clInfo Is Nothing Then
                End If

                If AddCPReadingsTerm And CPReadingsTermQID <> 0 Then
                    clInfo = Nothing
                    clInfo = oInspection.ChecklistMasterCollection.Item(CPReadingsTermQID)
                    If Not clInfo Is Nothing Then
                        If clInfo.Show Then
                            slTermPipeID = New SortedList
                            For Each nPipe As Integer In clInfo.PipeTermArrayList.ToArray
                                slTermPipeID.Add(npipe, htPipeIDIndex.Item(npipe))
                                alTerm1.Add(nPipe)
                                'alTerm2.Add(nPipe)
                                'alTerm3.Add(nPipe)
                                slTerm1.Add(htPipeIDIndex.Item(npipe), npipe)
                                'slTerm2.Add(htPipeIDIndex.Item(npipe), npipe)
                                'slTerm3.Add(htPipeIDIndex.Item(npipe), npipe)
                            Next
                            If hasResponses Then
                                For Each cp As MUSTER.Info.InspectionCPReadingsInfo In oInspection.CPReadingsCollection.Values
                                    If cp.QuestionID = CPReadingsTermQID Then
                                        If cp.RemoteReferCellPlacement Then
                                            cp.Deleted = False
                                            bolAddTermRemote = False
                                        ElseIf cp.GalvanicIC Then
                                            bolAddTermGalvanic = False
                                            cp.Deleted = False
                                        ElseIf cp.TestedByInspector Then
                                            nTermInspectorTestedID = cp.ID
                                            bolAddTermInspectorTested = False
                                            cp.Deleted = False
                                        Else
                                            If slTerm1.Contains(cp.TankPipeIndex) And cp.TankPipeEntityID = 10 Then
                                                If cp.Deleted Then
                                                    slTerm1.Remove(cp.TankPipeIndex)
                                                Else
                                                    'cp.Deleted = False
                                                    slTerm1.Remove(cp.TankPipeIndex)
                                                    If cp.PassFailIncon = 0 Or cp.PassFailIncon = 2 Then
                                                        If Not slTermWithResponseNo.Contains(cp.TankPipeIndex) Then
                                                            slTermWithResponseNo.Add(cp.TankPipeIndex, cp.TankPipeIndex)
                                                        End If
                                                    End If
                                                End If
                                                'ElseIf slTerm2.Contains(cp.TankPipeIndex) And cp.TankPipeEntityID = 10 Then
                                                '    cp.Deleted = False
                                                '    slTerm2.Remove(cp.TankPipeIndex)
                                                '    If cp.PassFailIncon = 0 Or cp.PassFailIncon = 2 Then
                                                '        If Not slTermWithResponseNo.Contains(cp.TankPipeIndex) Then
                                                '            slTermWithResponseNo.Add(cp.TankPipeIndex, cp.TankPipeIndex)
                                                '        End If
                                                '    End If
                                                'ElseIf slTerm3.Contains(cp.TankPipeIndex) And cp.TankPipeEntityID = 10 Then
                                                '    cp.Deleted = False
                                                '    slTerm3.Remove(cp.TankPipeIndex)
                                                '    If cp.PassFailIncon = 0 Or cp.PassFailIncon = 2 Then
                                                '        If Not slTermWithResponseNo.Contains(cp.TankPipeIndex) Then
                                                '            slTermWithResponseNo.Add(cp.TankPipeIndex, cp.TankPipeIndex)
                                                '        End If
                                                '    End If
                                            ElseIf clInfo.PipeTermArrayList.Contains(cp.TankPipeID) And cp.TankPipeEntityID = 10 Then
                                                If Not cp.Deleted Then
                                                    'cp.Deleted = False
                                                    If Not slTermPipeID.Contains(cp.TankPipeID) Then
                                                        slTermPipeID.Add(cp.TankPipeID, htPipeIDIndex.Item(cp.TankPipeID))
                                                    End If
                                                    If cp.PassFailIncon = 0 Or cp.PassFailIncon = 2 Then
                                                        If Not slTermWithResponseNo.Contains(cp.TankPipeIndex) Then
                                                            slTermWithResponseNo.Add(cp.TankPipeIndex, cp.TankPipeIndex)
                                                        End If
                                                    End If
                                                End If
                                            ElseIf Not [readOnly] Then
                                                cp.Deleted = True
                                            End If
                                            'If alTerm1.Contains(cp.TankPipeID) And cp.TankPipeEntityID = 10 Then
                                            '    cp.Deleted = False
                                            '    alTerm1.Remove(cp.TankPipeID)
                                            '    If cp.PassFailIncon = 0 Or cp.PassFailIncon = 2 Then
                                            '        If Not slTermWithResponseNo.Contains(cp.TankPipeIndex) Then
                                            '            slTermWithResponseNo.Add(cp.TankPipeIndex, cp.TankPipeIndex)
                                            '        End If
                                            '    End If
                                            'ElseIf alTerm2.Contains(cp.TankPipeID) And cp.TankPipeEntityID = 10 Then
                                            '    cp.Deleted = False
                                            '    alTerm2.Remove(cp.TankPipeID)
                                            '    If cp.PassFailIncon = 0 Or cp.PassFailIncon = 2 Then
                                            '        If Not slTermWithResponseNo.Contains(cp.TankPipeIndex) Then
                                            '            slTermWithResponseNo.Add(cp.TankPipeIndex, cp.TankPipeIndex)
                                            '        End If
                                            '    End If
                                            'ElseIf alTerm3.Contains(cp.TankPipeID) And cp.TankPipeEntityID = 10 Then
                                            '    cp.Deleted = False
                                            '    alTerm3.Remove(cp.TankPipeID)
                                            '    If cp.PassFailIncon = 0 Or cp.PassFailIncon = 2 Then
                                            '        If Not slTermWithResponseNo.Contains(cp.TankPipeIndex) Then
                                            '            slTermWithResponseNo.Add(cp.TankPipeIndex, cp.TankPipeIndex)
                                            '        End If
                                            '    End If
                                            'ElseIf clInfo.PipeTermArrayList.Contains(cp.TankPipeID) And cp.TankPipeEntityID = 10 Then
                                            '    cp.Deleted = False
                                            '    slTermPipeID.Add(cp.TankPipeID, htPipeIDIndex.Item(cp.TankPipeID))
                                            '    If cp.PassFailIncon = 0 Or cp.PassFailIncon = 2 Then
                                            '        If Not slTermWithResponseNo.Contains(cp.TankPipeIndex) Then
                                            '            slTermWithResponseNo.Add(cp.TankPipeIndex, cp.TankPipeIndex)
                                            '        End If
                                            '    End If
                                            'End If
                                        End If
                                    End If
                                Next
                            End If ' If hasResponses Then

                            If bolAddTermInspectorTested And Not [readOnly] Then
                                cpreading = New MUSTER.Info.InspectionCPReadingsInfo(0, _
                                            oInspection.ID, _
                                            CPReadingsTermQID, _
                                            0, _
                                            0, _
                                            0, _
                                            String.Empty, _
                                            String.Empty, _
                                            String.Empty, _
                                            String.Empty, _
                                            -1, _
                                            False, _
                                            String.Empty, _
                                            CDate("01/01/0001"), _
                                            String.Empty, _
                                            CDate("01/01/0001"), _
                                            -1, _
                                            False, _
                                            False, _
                                            -1, True, True)
                                oInspectionCPReadings.Add(cpreading)
                            ElseIf Not bolAddTermInspectorTested And Not [readOnly] Then
                                If nTermInspectorTestedID <> 0 Then
                                    If Not oInspection.CPReadingsCollection.Item(nTermInspectorTestedID) Is Nothing Then
                                        If oInspection.CPReadingsCollection.Item(nTermInspectorTestedID).TestedByInspectorResponse = False Then
                                            bolTermInspectorTestedResponse = False
                                        End If
                                    End If
                                End If
                            End If

                            If bolAddTermRemote And Not [readOnly] Then
                                cpreading = New MUSTER.Info.InspectionCPReadingsInfo(0, _
                                            oInspection.ID, _
                                            CPReadingsTermQID, _
                                            0, _
                                            0, _
                                            0, _
                                            String.Empty, _
                                            String.Empty, _
                                            String.Empty, _
                                            String.Empty, _
                                            -1, _
                                            False, _
                                            String.Empty, _
                                            CDate("01/01/0001"), _
                                            String.Empty, _
                                            CDate("01/01/0001"), _
                                            -1, _
                                            True, _
                                            False, _
                                            -1, False, False)
                                oInspectionCPReadings.Add(cpreading)
                            End If

                            If bolAddTermGalvanic And Not [readOnly] Then
                                cpreading = New MUSTER.Info.InspectionCPReadingsInfo(0, _
                                            oInspection.ID, _
                                            CPReadingsTermQID, _
                                            0, _
                                            0, _
                                            0, _
                                            String.Empty, _
                                            String.Empty, _
                                            String.Empty, _
                                            String.Empty, _
                                            -1, _
                                            False, _
                                            String.Empty, _
                                            CDate("01/01/0001"), _
                                            String.Empty, _
                                            CDate("01/01/0001"), _
                                            0, _
                                            False, _
                                            True, _
                                            -1, False, False)
                                oInspectionCPReadings.Add(cpreading)
                            End If

                            If slTerm1.Count > 0 And Not [readOnly] Then
                                'If (slTerm1.Count > 0 Or slTerm2.Count > 0 Or slTerm3.Count > 0) And Not [readOnly] Then
                                Dim slTerm As New SortedList
                                Dim addTermNos As Integer = 0
                                For i As Integer = 0 To slTerm1.Count - 1
                                    slTerm.Add(slTerm1.GetKey(i), 1)
                                Next
                                'For i As Integer = 0 To slTerm2.Count - 1
                                '    If slTerm.Contains(slTerm2.GetKey(i)) Then
                                '        slTerm.Item(slTerm2.GetKey(i)) += 1
                                '    Else
                                '        slTerm.Add(slTerm2.GetKey(i), 1)
                                '    End If
                                'Next
                                'For i As Integer = 0 To slTerm3.Count - 1
                                '    If slTerm.Contains(slTerm3.GetKey(i)) Then
                                '        slTerm.Item(slTerm3.GetKey(i)) += 1
                                '    Else
                                '        slTerm.Add(slTerm3.GetKey(i), 1)
                                '    End If
                                'Next
                                For i As Integer = 0 To slTerm.Count - 1
                                    addTermNos = slTerm.GetByIndex(i)
                                    For j As Integer = 0 To addTermNos - 1
                                        cpreading = New MUSTER.Info.InspectionCPReadingsInfo(0, _
                                                    oInspection.ID, _
                                                    CPReadingsTermQID, _
                                                    slTerm1.Item(slTerm.GetKey(i)), _
                                                    slTerm.GetKey(i), _
                                                    10, _
                                                    String.Empty, _
                                                    String.Empty, _
                                                    String.Empty, _
                                                    String.Empty, _
                                                    -1, _
                                                    False, _
                                                    String.Empty, _
                                                    CDate("01/01/0001"), _
                                                    String.Empty, _
                                                    CDate("01/01/0001"), _
                                                    nMaxTermCPLineNum, _
                                                    False, _
                                                    False, _
                                                    -1, False, False)
                                        oInspectionCPReadings.Add(cpreading)
                                        nMaxTermCPLineNum += 1
                                    Next
                                Next
                            End If ' If (slPipe1.Count > 0 Or slPipe2.Count > 0) And Not [readOnly] Then
                            'If alTerm1.Count > 0 And Not [readOnly] Then
                            '    For Each nTerm As Integer In alTerm1.ToArray
                            '        cpreading = New MUSTER.Info.InspectionCPReadingsInfo(0, _
                            '                    oInspection.ID, _
                            '                    CPReadingsTermQID, _
                            '                    nTerm, _
                            '                    slTermPipeID.Item(nTerm), _
                            '                    10, _
                            '                    String.Empty, _
                            '                    String.Empty, _
                            '                    String.Empty, _
                            '                    String.Empty, _
                            '                    -1, _
                            '                    False, _
                            '                    String.Empty, _
                            '                    CDate("01/01/0001"), _
                            '                    String.Empty, _
                            '                    CDate("01/01/0001"), _
                            '                    nMaxTermCPLineNum, _
                            '                    False, _
                            '                    False, _
                            '                    -1)
                            '        oInspectionCPReadings.Add(cpreading)
                            '        nMaxTermCPLineNum += 1
                            '    Next
                            'End If ' If alTerm1.Count > 0 And Not [readOnly] Then
                            'If alTerm2.Count > 0 And Not [readOnly] Then
                            '    For Each nTerm As Integer In alTerm2.ToArray
                            '        cpreading = New MUSTER.Info.InspectionCPReadingsInfo(0, _
                            '                    oInspection.ID, _
                            '                    CPReadingsTermQID, _
                            '                    nTerm, _
                            '                    slTermPipeID.Item(nTerm), _
                            '                    10, _
                            '                    String.Empty, _
                            '                    String.Empty, _
                            '                    String.Empty, _
                            '                    String.Empty, _
                            '                    -1, _
                            '                    False, _
                            '                    String.Empty, _
                            '                    CDate("01/01/0001"), _
                            '                    String.Empty, _
                            '                    CDate("01/01/0001"), _
                            '                    nMaxTermCPLineNum, _
                            '                    False, _
                            '                    False, _
                            '                    -1)
                            '        oInspectionCPReadings.Add(cpreading)
                            '        nMaxTermCPLineNum += 1
                            '    Next
                            'End If ' If alTerm2.Count > 0 And Not [readOnly] Then
                            'If alTerm3.Count > 0 And Not [readOnly] Then
                            '    For Each nTerm As Integer In alTerm3.ToArray
                            '        cpreading = New MUSTER.Info.InspectionCPReadingsInfo(0, _
                            '                    oInspection.ID, _
                            '                    CPReadingsTermQID, _
                            '                    nTerm, _
                            '                    slTermPipeID.Item(nTerm), _
                            '                    10, _
                            '                    String.Empty, _
                            '                    String.Empty, _
                            '                    String.Empty, _
                            '                    String.Empty, _
                            '                    -1, _
                            '                    False, _
                            '                    String.Empty, _
                            '                    CDate("01/01/0001"), _
                            '                    String.Empty, _
                            '                    CDate("01/01/0001"), _
                            '                    nMaxTermCPLineNum, _
                            '                    False, _
                            '                    False, _
                            '                    -1)
                            '        oInspectionCPReadings.Add(cpreading)
                            '        nMaxTermCPLineNum += 1
                            '    Next
                            'End If ' If alTerm3.Count > 0 And Not [readOnly] Then

                            ' if there were any cp readings with response no, check if citation exists and mark deleted
                            ' else create citation
                            If slTermWithResponseNo.Count > 0 Then
                                Dim citation As MUSTER.Info.InspectionCitationInfo
                                If bolTermInspectorTestedResponse Then
                                    Dim createCitation As Boolean = True
                                    Dim strTerms As String = String.Empty

                                    For i As Integer = 0 To slTermWithResponseNo.Count - 1
                                        strTerms += "T" + slTermWithResponseNo.GetByIndex(i).ToString + ", "
                                    Next
                                    If strTerms <> String.Empty Then
                                        strTerms = strTerms.Trim.TrimEnd(",")
                                    End If

                                    For Each Citation In oInspection.CitationsCollection.Values
                                        If Citation.QuestionID = CPReadingsTermQID Then
                                            Citation.Deleted = False
                                            createCitation = False
                                            Citation.CCAT = strTerms
                                            Exit For
                                        End If
                                    Next
                                    If createCitation Then
                                        citation = New MUSTER.Info.InspectionCitationInfo(0, _
                                        oInspection.ID, _
                                        CPReadingsTermQID, _
                                        oInspection.FacilityID, _
                                        0, _
                                        0, _
                                        clInfo.Citation, _
                                        strTerms, _
                                        False, _
                                        CDate("01/01/0001"), _
                                        CDate("01/01/0001"), _
                                        CDate("01/01/0001"), _
                                        False, _
                                        String.Empty, _
                                        CDate("01/01/0001"), _
                                        String.Empty, _
                                        CDate("01/01/0001"))
                                        oInspectionCitation.InspectionInfo = oInspection
                                        oInspectionCitation.Add(citation)
                                    End If ' If createCitation Then
                                Else
                                    ' mark citations deleted
                                    For Each Citation In oInspection.CitationsCollection.Values
                                        If Citation.QuestionID = CPReadingsTermQID Then
                                            Citation.Deleted = True
                                            Exit For
                                        End If
                                    Next
                                    For Each discrep As MUSTER.Info.InspectionDiscrepInfo In oInspection.DiscrepsCollection.Values
                                        If discrep.QuestionID = CPReadingsTermQID Then
                                            discrep.Deleted = True
                                            Exit For
                                        End If
                                    Next
                                End If
                            End If ' If slTermWithResponseNo.Count > 0 Then

                        End If ' If clInfo.Show Then
                    End If ' If Not clInfo Is Nothing Then
                End If

                'If AddCPReadingsTank Or AddCPReadingsPipe Or AddCPReadingsTerm Then
                '    For Each tnk In oOwner.Facility.TankCollection.Values
                '        bolAddTankCP = True
                '        If tnk.TankStatus = 424 Or tnk.TankStatus = 429 Then ' CIU / TOSI
                '            If Not slTankID.Contains(tnk.TankId) Then
                '                slTankID.Add(tnk.TankId, tnk.TankIndex)
                '            End If
                '            If hasResponses Then
                '                For Each cp As MUSTER.Info.InspectionCPReadingsInfo In oInspection.CPReadingsCollection.Values
                '                    If cp.QuestionID = CPReadingsTankQID And cp.TankPipeID = tnk.TankId And cp.TankPipeEntityID = 12 Then
                '                        cp.Deleted = False
                '                        bolAddTankCP = False
                '                    End If
                '                Next
                '            End If
                '            If AddCPReadingsTank And bolAddTankCP And Not [readOnly] Then
                '                cpreading = New MUSTER.Info.InspectionCPReadingsInfo(0, _
                '                            oInspection.ID, _
                '                            CPReadingsTankQID, _
                '                            tnk.TankId, _
                '                            tnk.TankIndex, _
                '                            12, _
                '                            False, _
                '                            False, _
                '                            False, _
                '                            String.Empty, _
                '                            String.Empty, _
                '                            String.Empty, _
                '                            String.Empty, _
                '                            -1, _
                '                            False, _
                '                            String.Empty, _
                '                            CDate("01/01/0001"), _
                '                            String.Empty, _
                '                            CDate("01/01/0001"), _
                '                            nMaxTankCPLineNum)
                '                oInspectionCPReadings.Add(cpreading)
                '                nMaxTankCPLineNum += 1
                '            End If
                '        End If

                '        If AddCPReadingsPipe Or AddCPReadingsTerm Then
                '            ' pipe / term
                '            For Each pipe In tnk.pipesCollection.Values
                '                bolAddPipeCP = True
                '                bolAddTermCP = True
                '                If pipe.PipeStatusDesc = 424 Or pipe.PipeStatusDesc = 429 Then ' CIU / TOSI
                '                    If Not slPipeID.Contains(pipe.PipeID) Then
                '                        slPipeID.Add(pipe.PipeID, pipe.Index)
                '                    End If
                '                    If Not slTermPipeID.Contains(pipe.PipeID) Then
                '                        slTermPipeID.Add(pipe.PipeID, pipe.Index)
                '                    End If
                '                    If hasResponses Then
                '                        For Each cp As MUSTER.Info.InspectionCPReadingsInfo In oInspection.CPReadingsCollection.Values
                '                            If cp.QuestionID = CPReadingsPipeQID And cp.TankPipeID = pipe.PipeID And cp.TankPipeEntityID = 10 Then
                '                                cp.Deleted = False
                '                                bolAddPipeCP = False
                '                            ElseIf cp.QuestionID = CPReadingsTermQID And cp.TankPipeID = pipe.PipeID And cp.TankPipeEntityID = 10 Then
                '                                cp.Deleted = False
                '                                bolAddTermCP = False
                '                            End If
                '                        Next
                '                    End If
                '                    ' pipe cp readings
                '                    If AddCPReadingsPipe And bolAddPipeCP And Not [readOnly] Then
                '                        cpreading = New MUSTER.Info.InspectionCPReadingsInfo(0, _
                '                                    oInspection.ID, _
                '                                    CPReadingsPipeQID, _
                '                                    pipe.PipeID, _
                '                                    pipe.Index, _
                '                                    10, _
                '                                    False, _
                '                                    False, _
                '                                    False, _
                '                                    String.Empty, _
                '                                    String.Empty, _
                '                                    String.Empty, _
                '                                    String.Empty, _
                '                                    -1, _
                '                                    False, _
                '                                    String.Empty, _
                '                                    CDate("01/01/0001"), _
                '                                    String.Empty, _
                '                                    CDate("01/01/0001"), _
                '                                    nMaxPipeCPLineNum)
                '                        oInspectionCPReadings.Add(cpreading)
                '                        nMaxPipeCPLineNum += 1

                '                        cpreading = New MUSTER.Info.InspectionCPReadingsInfo(0, _
                '                                    oInspection.ID, _
                '                                    CPReadingsPipeQID, _
                '                                    pipe.PipeID, _
                '                                    pipe.Index, _
                '                                    10, _
                '                                    False, _
                '                                    False, _
                '                                    False, _
                '                                    String.Empty, _
                '                                    String.Empty, _
                '                                    String.Empty, _
                '                                    String.Empty, _
                '                                    -1, _
                '                                    False, _
                '                                    String.Empty, _
                '                                    CDate("01/01/0001"), _
                '                                    String.Empty, _
                '                                    CDate("01/01/0001"), _
                '                                    nMaxPipeCPLineNum)
                '                        oInspectionCPReadings.Add(cpreading)
                '                        nMaxPipeCPLineNum += 1
                '                    End If

                '                    ' term cp readings
                '                    If AddCPReadingsTerm And bolAddTermCP And Not [readOnly] Then
                '                        cpreading = New MUSTER.Info.InspectionCPReadingsInfo(0, _
                '                                    oInspection.ID, _
                '                                    CPReadingsTermQID, _
                '                                    pipe.PipeID, _
                '                                    pipe.Index, _
                '                                    10, _
                '                                    False, _
                '                                    False, _
                '                                    False, _
                '                                    String.Empty, _
                '                                    String.Empty, _
                '                                    String.Empty, _
                '                                    String.Empty, _
                '                                    -1, _
                '                                    False, _
                '                                    String.Empty, _
                '                                    CDate("01/01/0001"), _
                '                                    String.Empty, _
                '                                    CDate("01/01/0001"), _
                '                                    nMaxTermCPLineNum)
                '                        oInspectionCPReadings.Add(cpreading)
                '                        nMaxTermCPLineNum += 1

                '                        cpreading = New MUSTER.Info.InspectionCPReadingsInfo(0, _
                '                                    oInspection.ID, _
                '                                    CPReadingsTermQID, _
                '                                    pipe.PipeID, _
                '                                    pipe.Index, _
                '                                    10, _
                '                                    False, _
                '                                    False, _
                '                                    False, _
                '                                    String.Empty, _
                '                                    String.Empty, _
                '                                    String.Empty, _
                '                                    String.Empty, _
                '                                    -1, _
                '                                    False, _
                '                                    String.Empty, _
                '                                    CDate("01/01/0001"), _
                '                                    String.Empty, _
                '                                    CDate("01/01/0001"), _
                '                                    nMaxTermCPLineNum)
                '                        oInspectionCPReadings.Add(cpreading)
                '                        nMaxTermCPLineNum += 1

                '                        cpreading = New MUSTER.Info.InspectionCPReadingsInfo(0, _
                '                                    oInspection.ID, _
                '                                    CPReadingsTermQID, _
                '                                    pipe.PipeID, _
                '                                    pipe.Index, _
                '                                    10, _
                '                                    False, _
                '                                    False, _
                '                                    False, _
                '                                    String.Empty, _
                '                                    String.Empty, _
                '                                    String.Empty, _
                '                                    String.Empty, _
                '                                    -1, _
                '                                    False, _
                '                                    String.Empty, _
                '                                    CDate("01/01/0001"), _
                '                                    String.Empty, _
                '                                    CDate("01/01/0001"), _
                '                                    nMaxTermCPLineNum)
                '                        oInspectionCPReadings.Add(cpreading)
                '                        nMaxTermCPLineNum += 1
                '                    End If
                '                End If ' if pipe = ciu / tosi
                '            Next ' for each pipe
                '        End If ' If AddCPReadingsPipe Or AddCPReadingsTerm Then
                '    Next ' for each tank
                'End If ' If AddCPReadingsTank Or AddCPReadingsPipe Or AddCPReadingsTerm Then
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Private Sub AddMonitorWells(ByVal hasResponses As Boolean, Optional ByVal [readOnly] As Boolean = False)
            Dim bolAddMWTank, bolAddMWPipe As Boolean
            Dim slTankWithResponseNo As New SortedList
            Dim slPipeWithResponseNo As New SortedList
            Dim mwells As MUSTER.Info.InspectionMonitorWellsInfo
            Dim clInfo As MUSTER.Info.InspectionChecklistMasterInfo

            oInspectionMonitorWells.InspectionInfo = oInspection

            bolAddMWTank = True
            bolAddMWPipe = True

            Try
                ' tank m wells
                clInfo = oInspection.ChecklistMasterCollection.Item(MonitorWellsTankQID)
                If Not clInfo Is Nothing Then
                    If clInfo.Show Then
                        If hasResponses Then
                            For Each mwells In oInspection.MonitorWellsCollection.Values
                                If Math.Abs(IIf(mwells.QuestionID < -100000, mwells.QuestionID + 100000, mwells.QuestionID)) = MonitorWellsTankQID Then
                                    bolAddMWTank = False
                                    mwells.Deleted = False
                                    If (mwells.SurfaceSealed = 0 Or mwells.WellCaps = 0) Then ' mwells.WellNumber <> 0 And 
                                        If Not slTankWithResponseNo.Contains(mwells.WellNumber) Then
                                            slTankWithResponseNo.Add(mwells.WellNumber, mwells.WellNumber)
                                        End If
                                    End If
                                End If
                            Next
                        End If ' If hasResponses Then

                        ' tank
                        If AddMonitorWellsTank And bolAddMWTank And Not [readOnly] Then
                            mwells = New MUSTER.Info.InspectionMonitorWellsInfo(0, _
                                    oInspection.ID, _
                                    MonitorWellsTankQID, _
                                    True, _
                                    0, _
                                    String.Empty, _
                                    String.Empty, _
                                    String.Empty, _
                                    -1, _
                                    -1, _
                                    String.Empty, _
                                    False, _
                                    String.Empty, _
                                    CDate("01/01/0001"), _
                                    String.Empty, _
                                    CDate("01/01/0001"), _
                                    nMaxTankBedMWLineNum)
                            oInspectionMonitorWells.Add(mwells)
                            nMaxTankBedMWLineNum += 1

                            mwells = New MUSTER.Info.InspectionMonitorWellsInfo(0, _
                                    oInspection.ID, _
                                    MonitorWellsTankQID, _
                                    True, _
                                    0, _
                                    String.Empty, _
                                    String.Empty, _
                                    String.Empty, _
                                    -1, _
                                    -1, _
                                    String.Empty, _
                                    False, _
                                    String.Empty, _
                                    CDate("01/01/0001"), _
                                    String.Empty, _
                                    CDate("01/01/0001"), _
                                    nMaxTankBedMWLineNum)
                            oInspectionMonitorWells.Add(mwells)
                            nMaxTankBedMWLineNum += 1

                            mwells = New MUSTER.Info.InspectionMonitorWellsInfo(0, _
                                    oInspection.ID, _
                                    MonitorWellsTankQID, _
                                    True, _
                                    0, _
                                    String.Empty, _
                                    String.Empty, _
                                    String.Empty, _
                                    -1, _
                                    -1, _
                                    String.Empty, _
                                    False, _
                                    String.Empty, _
                                    CDate("01/01/0001"), _
                                    String.Empty, _
                                    CDate("01/01/0001"), _
                                    nMaxTankBedMWLineNum)
                            oInspectionMonitorWells.Add(mwells)
                            nMaxTankBedMWLineNum += 1

                            mwells = New MUSTER.Info.InspectionMonitorWellsInfo(0, _
                                    oInspection.ID, _
                                    MonitorWellsTankQID, _
                                    True, _
                                    0, _
                                    String.Empty, _
                                    String.Empty, _
                                    String.Empty, _
                                    -1, _
                                    -1, _
                                    String.Empty, _
                                    False, _
                                    String.Empty, _
                                    CDate("01/01/0001"), _
                                    String.Empty, _
                                    CDate("01/01/0001"), _
                                    nMaxTankBedMWLineNum)
                            oInspectionMonitorWells.Add(mwells)
                            nMaxTankBedMWLineNum += 1
                        End If ' If AddMonitorWellsTank And bolAddMWTank And Not [readOnly] Then

                        ' if there were any monitor wells with response no, check if citation exists and mark deleted
                        ' else create citation
                        If slTankWithResponseNo.Count > 0 Then
                            Dim createCitation As Boolean = False
                            Dim citation As MUSTER.Info.InspectionCitationInfo
                            Dim strTanks As String = String.Empty

                            For i As Integer = 0 To slTankWithResponseNo.Count - 1
                                If slTankWithResponseNo.GetByIndex(i) <> 0 Then
                                    strTanks += slTankWithResponseNo.GetByIndex(i).ToString + ", "
                                End If
                            Next
                            If strTanks <> String.Empty Then
                                strTanks = strTanks.Trim.TrimEnd(",")
                            End If

                            For Each citation In oInspection.CitationsCollection.Values
                                If citation.QuestionID = MonitorWellsTankQID Then
                                    citation.Deleted = False
                                    createCitation = False
                                    citation.CCAT = strTanks
                                    Exit For
                                End If
                            Next
                            If createCitation Then
                                citation = New MUSTER.Info.InspectionCitationInfo(0, _
                                oInspection.ID, _
                                MonitorWellsTankQID, _
                                oInspection.FacilityID, _
                                0, _
                                0, _
                                clInfo.Citation, _
                                strTanks, _
                                False, _
                                CDate("01/01/0001"), _
                                CDate("01/01/0001"), _
                                CDate("01/01/0001"), _
                                False, _
                                String.Empty, _
                                CDate("01/01/0001"), _
                                String.Empty, _
                                CDate("01/01/0001"))
                                oInspectionCitation.InspectionInfo = oInspection
                                oInspectionCitation.Add(citation)
                            End If ' If createCitation Then
                        End If ' If slTankWithResponseNo.Count > 0 Then

                    End If ' If clInfo.Show Then
                End If ' If Not clInfo Is Nothing Then

                ' pipe m wells
                clInfo = oInspection.ChecklistMasterCollection.Item(MonitorWellsPipeQID)
                If Not clInfo Is Nothing Then
                    If clInfo.Show Then
                        If hasResponses Then
                            For Each mwells In oInspection.MonitorWellsCollection.Values
                                If Math.Abs(IIf(mwells.QuestionID < -100000, mwells.QuestionID + 100000, mwells.QuestionID)) = MonitorWellsPipeQID Then
                                    bolAddMWPipe = False
                                    mwells.Deleted = False
                                    If (mwells.SurfaceSealed = 0 Or mwells.WellCaps = 0) Then ' mwells.WellNumber <> 0 And 
                                        If Not slPipeWithResponseNo.Contains(mwells.WellNumber) Then
                                            slPipeWithResponseNo.Add(mwells.WellNumber, mwells.WellNumber)
                                        End If
                                    End If
                                End If
                            Next
                        End If ' If hasResponses Then

                        ' pipe
                        If AddMonitorWellsPipe And bolAddMWPipe And Not [readOnly] Then
                            mwells = New MUSTER.Info.InspectionMonitorWellsInfo(0, _
                                    oInspection.ID, _
                                    MonitorWellsPipeQID, _
                                    False, _
                                    0, _
                                    String.Empty, _
                                    String.Empty, _
                                    String.Empty, _
                                    -1, _
                                    -1, _
                                    String.Empty, _
                                    False, _
                                    String.Empty, _
                                    CDate("01/01/0001"), _
                                    String.Empty, _
                                    CDate("01/01/0001"), _
                                    nMaxLineMWLineNum)
                            oInspectionMonitorWells.Add(mwells)
                            nMaxLineMWLineNum += 1
                        End If ' If AddMonitorWellsPipe And bolAddMWPipe And Not [readOnly] Then

                        ' if there were any monitor wells with response no, check if citation exists and mark deleted
                        ' else create citation
                        If slPipeWithResponseNo.Count > 0 Then
                            Dim createCitation As Boolean = False
                            Dim citation As MUSTER.Info.InspectionCitationInfo
                            Dim strPipes As String = String.Empty

                            For i As Integer = 0 To slPipeWithResponseNo.Count - 1
                                If slPipeWithResponseNo.GetByIndex(i) <> 0 Then
                                    strPipes += slPipeWithResponseNo.GetByIndex(i).ToString + ", "
                                End If
                            Next
                            If strPipes <> String.Empty Then
                                strPipes = strPipes.Trim.TrimEnd(",")
                            End If

                            For Each citation In oInspection.CitationsCollection.Values
                                If citation.QuestionID = MonitorWellsPipeQID Then
                                    citation.Deleted = False
                                    createCitation = False
                                    citation.CCAT = strPipes
                                    Exit For
                                End If
                            Next
                            If createCitation Then
                                citation = New MUSTER.Info.InspectionCitationInfo(0, _
                                oInspection.ID, _
                                MonitorWellsPipeQID, _
                                oInspection.FacilityID, _
                                0, _
                                0, _
                                clInfo.Citation, _
                                strPipes, _
                                False, _
                                CDate("01/01/0001"), _
                                CDate("01/01/0001"), _
                                CDate("01/01/0001"), _
                                False, _
                                String.Empty, _
                                CDate("01/01/0001"), _
                                String.Empty, _
                                CDate("01/01/0001"))
                                oInspectionCitation.InspectionInfo = oInspection
                                oInspectionCitation.Add(citation)
                            End If ' If createCitation Then
                        End If ' If slPipeWithResponseNo.Count > 0 Then

                    End If ' If clInfo.Show Then
                End If ' If Not clInfo Is Nothing Then

                ' #2807
                ' tank pipe m wells
                If AddMonitorWellsTank = False Or AddMonitorWellsPipe = False Then
                    clInfo = oInspection.ChecklistMasterCollection.Item(MonitorWellsTankPipeQID)
                    If Not clInfo Is Nothing Then
                        If clInfo.Show Then
                            'If hasResponses Then
                            For Each mwells In oInspection.MonitorWellsCollection.Values
                                If Math.Abs(IIf(mwells.QuestionID < -100000, mwells.QuestionID + 100000, mwells.QuestionID)) = MonitorWellsTankPipeQID Then
                                    mwells.Deleted = False
                                End If
                            Next
                            'End If ' If hasResponses Then
                        End If ' If clInfo.Show Then
                    End If ' If Not clInfo Is Nothing Then
                End If


                'If hasResponses Then
                '    For Each mwells In oInspection.MonitorWellsCollection.Values
                '        If mwells.QuestionID = MonitorWellsTankQID Then
                '            bolAddMWTank = False
                '            mwells.Deleted = False
                '            If Not slTankWithResponseNo.Contains(mwells.WellNumber) Then
                '                slTankWithResponseNo.Add(mwells.WellNumber, mwells.WellNumber)
                '            End If
                '            'Exit For
                '        ElseIf mwells.QuestionID = MonitorWellsPipeQID Then
                '            bolAddMWPipe = False
                '            mwells.Deleted = False
                '            If Not slPipeWithResponseNo.Contains(mwells.WellNumber) Then
                '                slPipeWithResponseNo.Add(mwells.WellNumber, mwells.WellNumber)
                '            End If
                '            'Exit For
                '        End If
                '    Next
                'End If
                '' tank
                'If AddMonitorWellsTank And bolAddMWTank And Not [readOnly] Then
                '    mwells = New MUSTER.Info.InspectionMonitorWellsInfo(0, _
                '            oInspection.ID, _
                '            MonitorWellsTankQID, _
                '            False, _
                '            0, _
                '            0.0, _
                '            0.0, _
                '            0.0, _
                '            -1, _
                '            -1, _
                '            String.Empty, _
                '            False, _
                '            String.Empty, _
                '            CDate("01/01/0001"), _
                '            String.Empty, _
                '            CDate("01/01/0001"), _
                '            nMaxTankBedMWLineNum)
                '    oInspectionMonitorWells.Add(mwells)
                '    nMaxTankBedMWLineNum += 1

                '    mwells = New MUSTER.Info.InspectionMonitorWellsInfo(0, _
                '            oInspection.ID, _
                '            MonitorWellsTankQID, _
                '            False, _
                '            0, _
                '            0.0, _
                '            0.0, _
                '            0.0, _
                '            -1, _
                '            -1, _
                '            String.Empty, _
                '            False, _
                '            String.Empty, _
                '            CDate("01/01/0001"), _
                '            String.Empty, _
                '            CDate("01/01/0001"), _
                '            nMaxTankBedMWLineNum)
                '    oInspectionMonitorWells.Add(mwells)
                '    nMaxTankBedMWLineNum += 1

                '    mwells = New MUSTER.Info.InspectionMonitorWellsInfo(0, _
                '            oInspection.ID, _
                '            MonitorWellsTankQID, _
                '            False, _
                '            0, _
                '            0.0, _
                '            0.0, _
                '            0.0, _
                '            -1, _
                '            -1, _
                '            String.Empty, _
                '            False, _
                '            String.Empty, _
                '            CDate("01/01/0001"), _
                '            String.Empty, _
                '            CDate("01/01/0001"), _
                '            nMaxTankBedMWLineNum)
                '    oInspectionMonitorWells.Add(mwells)
                '    nMaxTankBedMWLineNum += 1

                '    mwells = New MUSTER.Info.InspectionMonitorWellsInfo(0, _
                '            oInspection.ID, _
                '            MonitorWellsTankQID, _
                '            False, _
                '            0, _
                '            0.0, _
                '            0.0, _
                '            0.0, _
                '            -1, _
                '            -1, _
                '            String.Empty, _
                '            False, _
                '            String.Empty, _
                '            CDate("01/01/0001"), _
                '            String.Empty, _
                '            CDate("01/01/0001"), _
                '            nMaxTankBedMWLineNum)
                '    oInspectionMonitorWells.Add(mwells)
                '    nMaxTankBedMWLineNum += 1
                'End If
                '' pipe
                'If AddMonitorWellsPipe And bolAddMWPipe And Not [readOnly] Then
                '    mwells = New MUSTER.Info.InspectionMonitorWellsInfo(0, _
                '            oInspection.ID, _
                '            MonitorWellsPipeQID, _
                '            False, _
                '            0, _
                '            0.0, _
                '            0.0, _
                '            0.0, _
                '            -1, _
                '            -1, _
                '            String.Empty, _
                '            False, _
                '            String.Empty, _
                '            CDate("01/01/0001"), _
                '            String.Empty, _
                '            CDate("01/01/0001"), _
                '            nMaxLineMWLineNum)
                '    oInspectionMonitorWells.Add(mwells)
                '    nMaxLineMWLineNum += 1
                'End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Private Sub AddSOC(ByVal hasResponses As Boolean, Optional ByVal [readOnly] As Boolean = False)
            Dim bolAddSOC As Boolean = True
            Try
                If hasResponses Then
                    If oInspection.SOCsCollection.Count > 0 Then
                        For Each soc As MUSTER.Info.InspectionSOCInfo In oInspection.SOCsCollection.Values
                            If soc.InspectionID = oInspection.ID Then
                                bolAddSOC = False
                                soc.Deleted = False
                            End If
                        Next
                    End If
                End If
                If bolAddSOC And Not [readOnly] Then
                    Dim soc As New MUSTER.Info.InspectionSOCInfo(0, _
                                    oInspection.ID, _
                                    -1, _
                                    String.Empty, _
                                    String.Empty, _
                                    -1, _
                                    String.Empty, _
                                    String.Empty, _
                                    -1, _
                                    False, _
                                    String.Empty, _
                                    CDate("01/01/0001"), _
                                    String.Empty, _
                                    CDate("01/01/0001"), _
                                    False)
                    oInspectionSOC.InspectionInfo = oInspection
                    oInspectionSOC.Add(soc)
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Private Sub AddCCAT(ByVal hasResponses As Boolean, Optional ByVal [readOnly] As Boolean = False)
            Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo
            Dim ccat As MUSTER.Info.InspectionCCATInfo
            Dim tankID As Integer
            Dim pipeID As Integer
            Dim t As New BusinessLogic.pCompartment
            Dim ti As New Info.TankInfo

            Dim bolAddCCATTank, bolAddCCATPipe, bolAddCCATTerm As Boolean
            Try
                For Each checkList In oInspection.ChecklistMasterCollection.Values
                    If checkList.Show And checkList.CCAT Then

                        ' tank
                        For Each tankID In checkList.TankArrayList

                            t.Retrieve(ti, tankID)

                            Dim list As New Collections.ArrayList

                            If Not t.CompartmentCollection Is Nothing Then

                                For Each item As Info.CompartmentInfo In t.CompartmentCollection.Values
                                    list.Add(item.COMPARTMENTNumber)
                                Next

                            Else
                                list.Add(0)

                            End If

                            For Each g As Integer In list

                                bolAddCCATTank = True
                                If hasResponses Then
                                    For Each ccat In oInspection.CCATsCollection.Values
                                        If ccat.TankPipeID = tankID And _
                                            ccat.TankPipeEntityID = 12 And _
                                            ccat.QuestionID = checkList.ID And IIf(ccat.CompartmentID = 0, 1, ccat.CompartmentID) = IIf(g = 0, 1, g) Then
                                            bolAddCCATTank = False
                                            ccat.Deleted = False
                                            'Exit For
                                        End If
                                    Next
                                End If
                                If bolAddCCATTank And Not [readOnly] Then
                                    ccat = New MUSTER.Info.InspectionCCATInfo(0, _
                                            oInspection.ID, _
                                            checkList.ID, _
                                            tankID, _
                                            12, _
                                            False, _
                                            False, _
                                            String.Empty, _
                                            False, _
                                            String.Empty, _
                                            CDate("01/01/0001"), _
                                            String.Empty, _
                                            CDate("01/01/0001"), IIf(g = 0, 1, g))
                                    oInspectionCCAT.InspectionInfo = oInspection
                                    oInspectionCCAT.Add(ccat)
                                End If





                            Next g

                        Next

                        ' pipe
                        For Each pipeID In checkList.PipeArrayList
                            bolAddCCATPipe = True
                            If hasResponses Then
                                For Each ccat In oInspection.CCATsCollection.Values
                                    If ccat.TankPipeID = pipeID And _
                                        ccat.TankPipeEntityID = 10 And _
                                        Not ccat.Termination And _
                                        ccat.QuestionID = checkList.ID Then

                                        bolAddCCATPipe = False
                                        ccat.Deleted = False
                                        ccat.CompartmentID = t.COMPARTMENTNumber
                                        'Exit For
                                    End If
                                Next
                            End If
                            If bolAddCCATPipe And Not [readOnly] Then
                                ccat = New MUSTER.Info.InspectionCCATInfo(0, _
                                        oInspection.ID, _
                                        checkList.ID, _
                                        pipeID, _
                                        10, _
                                        False, _
                                        False, _
                                        String.Empty, _
                                        False, _
                                        String.Empty, _
                                        CDate("01/01/0001"), _
                                        String.Empty, _
                                        CDate("01/01/0001"))
                                oInspectionCCAT.InspectionInfo = oInspection
                                oInspectionCCAT.Add(ccat)
                            End If
                        Next
                        ' term
                        For Each pipeID In checkList.PipeTermArrayList
                            bolAddCCATTerm = True
                            If hasResponses Then
                                For Each ccat In oInspection.CCATsCollection.Values
                                    If ccat.TankPipeID = pipeID And _
                                        ccat.TankPipeEntityID = 10 And _
                                        ccat.Termination And _
                                        ccat.QuestionID = checkList.ID Then
                                        bolAddCCATTerm = False
                                        ccat.Deleted = False
                                        'Exit For
                                    End If
                                Next
                            End If
                            If bolAddCCATTerm And Not [readOnly] Then
                                ccat = New MUSTER.Info.InspectionCCATInfo(0, _
                                        oInspection.ID, _
                                        checkList.ID, _
                                        pipeID, _
                                        10, _
                                        False, _
                                        True, _
                                        String.Empty, _
                                        False, _
                                        String.Empty, _
                                        CDate("01/01/0001"), _
                                        String.Empty, _
                                        CDate("01/01/0001"))
                                oInspectionCCAT.InspectionInfo = oInspection
                                oInspectionCCAT.Add(ccat)
                            End If
                        Next


                    End If
                Next

                t = Nothing

            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub


        Public Sub RefreshCCAT()

            oInspection.CCATsCollection.Clear()

            AddCCAT2(True, False)
        End Sub
        Private Sub AddCCAT2(ByVal hasResponses As Boolean, Optional ByVal [readOnly] As Boolean = False)
            Dim ccat As MUSTER.Info.InspectionCCATInfo
            Dim ds As DataTable


            Try

                If oInspectionCCAT Is Nothing Then
                    oInspectionCCAT = New BusinessLogic.pInspectionCCAT
                End If

                If oInspection.ID > 0 Then

                    ds = oInspectionCCAT.GetCCATTankPipeTermListForInspection(oInspection.ID)
                Else
                    ds = oInspectionCCAT.GetCCATTankPipeTermListForInspection(oInspection.ID, oInspection.FacilityID)
                End If


                oInspection.CCATsCollection.Clear()

                If Not ds Is Nothing AndAlso ds.Rows.Count > 0 Then

                    For Each dr As DataRow In ds.Rows

                        Dim add As Boolean = True

                        For Each ccat In oInspection.CCATsCollection.Values

                            If ccat.TankPipeID = Convert.ToInt32(dr("Entity").ToString.Substring(1)) And _
                                ccat.TankPipeEntityID = dr("Mode") And _
                                ccat.Termination = IIf(dr("CODE") = "TERM", True, False) And _
                                dr("CODE") <> "TERM" And _
                                ccat.QuestionID = dr("Question_ID") And IIf(ccat.CompartmentID = 0, 1, ccat.CompartmentID) = IIf(dr("COMPARTMENTID") = 0, 1, dr("COMPARTMENTID")) Then
                                ccat.Deleted = False
                                add = False
                                Exit For
                            End If

                        Next

                        If add Then
                            ccat = New MUSTER.Info.InspectionCCATInfo(IIf(dr("CODE") = "TERM", dr("uid").ToString.Substring(1), 0), _
                                    oInspection.ID, _
                                    dr("Question_ID"), _
                                    Convert.ToInt32(dr("Entity").ToString.Substring(1)), _
                                     dr("Mode"), _
                                    False, _
                                    IIf(dr("CODE") = "TERM", True, False), _
                                    String.Empty, _
                                    False, _
                                    String.Empty, _
                                    CDate("01/01/0001"), _
                                    String.Empty, _
                                    CDate("01/01/0001"), dr("COMPARTMENTID"))
                            oInspectionCCAT.InspectionInfo = oInspection
                            oInspectionCCAT.Add(ccat)
                        End If



                    Next
                End If


            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub



        Private Sub AddCitation(ByVal hasResponses As Boolean, Optional ByVal [readOnly] As Boolean = False)
            Dim citation As MUSTER.Info.InspectionCitationInfo
            Dim checklist As MUSTER.Info.InspectionChecklistMasterInfo
            Dim resp As MUSTER.Info.InspectionResponsesInfo
            Dim createCitation As Boolean
            Try
                For Each resp In oInspection.ResponsesCollection.Values
                    If Not resp.Deleted Then

                        Dim msg As String = resp.QuestionID

                        checklist = oInspection.ChecklistMasterCollection.Item(resp.QuestionID)
                        If Not checklist Is Nothing AndAlso (msg <> 56 And msg <> 94) Then
                            ' add citation only if response = no and checklist is visible and has valid citation num
                            If resp.Response = 0 And (checklist.Show Or checklist.ID < -1) And _
                                    checklist.CheckListItemNumber < "8" And _
                                        checklist.Citation <> -1 Then
                                createCitation = True
                                If hasResponses Then
                                    For Each citation In oInspection.CitationsCollection.Values
                                        If citation.QuestionID = checklist.ID Then
                                            createCitation = False
                                            citation.Deleted = False
                                            Exit For
                                        End If
                                    Next
                                End If

                                If createCitation And Not [readOnly] AndAlso (msg <> 56 And msg <> 94) Then
                                    citation = New MUSTER.Info.InspectionCitationInfo(0, _
                                    oInspection.ID, _
                                    checklist.ID, _
                                    oInspection.FacilityID, _
                                    0, _
                                    0, _
                                    checklist.Citation, _
                                    String.Empty, _
                                    False, _
                                    CDate("01/01/0001"), _
                                    CDate("01/01/0001"), _
                                    CDate("01/01/0001"), _
                                    False, _
                                    String.Empty, _
                                    CDate("01/01/0001"), _
                                    String.Empty, _
                                    CDate("01/01/0001"))
                                    oInspectionCitation.InspectionInfo = oInspection
                                    oInspectionCitation.Add(citation)
                                End If
                            End If
                        End If
                    End If
                Next
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Private Sub AddDiscrep(ByVal hasResponses As Boolean, Optional ByVal [readOnly] As Boolean = False)
            Dim discrep As MUSTER.Info.InspectionDiscrepInfo
            Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo
            Dim resp As MUSTER.Info.InspectionResponsesInfo
            Dim createDiscrep As Boolean = True
            Try
                For Each resp In oInspection.ResponsesCollection.Values
                    If Not resp.Deleted Then
                        Dim msg = resp.QuestionID
                        checkList = oInspection.ChecklistMasterCollection.Item(resp.QuestionID)
                        ' add discrep only if response is no and checklist is visible and checklist item has discrep text
                        If Not checkList Is Nothing AndAlso (msg <> 56 And msg <> 94) Then
                            If resp.Response = 0 And (checkList.Show Or checkList.ID < -1) And _
                                    checkList.DiscrepText <> String.Empty Then
                                createDiscrep = True
                                If hasResponses Then
                                    For Each discrep In oInspection.DiscrepsCollection.Values
                                        If discrep.QuestionID = resp.QuestionID Then
                                            createDiscrep = False
                                            discrep.Deleted = False
                                            Exit For
                                        End If
                                    Next
                                End If
                                If createDiscrep AndAlso (msg <> 56 And msg <> 94) And Not [readOnly] Then
                                    ' get the inspection citation id (primary key for citation in inspection_citation table)
                                    Dim inspCitID As Int64 = 0
                                    For Each citation As MUSTER.Info.InspectionCitationInfo In oInspection.CitationsCollection.Values
                                        If citation.QuestionID = resp.QuestionID Then
                                            inspCitID = citation.ID
                                            Exit For
                                        End If
                                    Next
                                    discrep = New MUSTER.Info.InspectionDiscrepInfo(0, _
                                    oInspection.ID, _
                                    checkList.ID, _
                                    checkList.DiscrepText, _
                                    False, _
                                    String.Empty, _
                                    CDate("01/01/0001"), _
                                    String.Empty, _
                                    CDate("01/01/0001"), _
                                    False, _
                                    CDate("01/01/0001"), _
                                    inspCitID)
                                    oInspectionDiscrep.InspectionInfo = oInspection
                                    oInspectionDiscrep.Add(discrep)
                                End If
                            End If
                        End If
                    End If
                Next
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Private Sub AddComments(ByVal hasResponses As Boolean, Optional ByVal [readOnly] As Boolean = False)
            Dim insComments As MUSTER.Info.InspectionCommentsInfo
            Dim addComments As Boolean = True
            Try
                If hasResponses Then
                    If oInspection.InspectionCommentsCollection.Count > 0 Then
                        addComments = False
                    End If
                End If
                If addComments And Not [readOnly] Then
                    insComments = New MUSTER.Info.InspectionCommentsInfo(0, _
                                    oInspection.ID, _
                                    String.Empty, _
                                    False, _
                                    String.Empty, _
                                    CDate("01/01/0001"), _
                                    String.Empty, _
                                    CDate("01/01/0001"))
                    oInspectionComments.Add(insComments)
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Private Sub AddTankFuelType()
            Try
                slTankFuelType = New SortedList
                Dim strSQL As String = "SELECT DISTINCT TANK.TANK_ID, A.PROPERTY_NAME " + _
                                        "FROM tblREG_TANK TANK LEFT OUTER JOIN tblREG_COMPARTMENTS COMP ON COMP.TANK_ID = TANK.TANK_ID AND COMP.DELETED = 0 " + _
                                        "LEFT OUTER JOIN tblSYS_PROPERTY_MASTER A ON COMP.FUEL_TYPE_ID = A.PROPERTY_ID AND A.PROPERTY_ACTIVE = 'YES' " + _
                                        "WHERE TANK.DELETED = 0 AND TANK.FACILITY_ID = " + oInspection.FacilityID.ToString
                Dim ds As DataSet = oInspectionChecklistMasterDB.DBGetDS(strSQL)
                If Not ds Is Nothing Then
                    For Each dr As DataRow In ds.Tables(0).Rows
                        If slTankFuelType.Contains(dr("TANK_ID")) Then
                            slTankFuelType.Item(dr("TANK_ID")) += ", " + IIf(dr("PROPERTY_NAME") Is DBNull.Value, "N/A", dr("PROPERTY_NAME"))
                        Else
                            slTankFuelType.Add(dr("TANK_ID"), IIf(dr("PROPERTY_NAME") Is DBNull.Value, "N/A", dr("PROPERTY_NAME")))
                        End If
                    Next
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Private Sub AddPipeFuelType()
            Try
                slPipeFuelType = New SortedList
                Dim strSQL As String = "SELECT DISTINCT PIPES.PIPE_ID, A.PROPERTY_NAME " + _
                                        "FROM tblREG_PIPE PIPES LEFT OUTER JOIN tblREG_COMPARTMENTS_PIPES COMPARTMENTS_PIPES ON COMPARTMENTS_PIPES.PIPE_ID = PIPES.PIPE_ID AND COMPARTMENTS_PIPES.DELETED = 0 " + _
                                        "LEFT OUTER JOIN tblREG_COMPARTMENTS COMP ON COMP.TANK_ID = COMPARTMENTS_PIPES.TANK_ID AND COMP.DELETED = 0 " + _
                                        "LEFT OUTER JOIN tblSYS_PROPERTY_MASTER A ON COMP.FUEL_TYPE_ID = A.PROPERTY_ID AND A.PROPERTY_ACTIVE = 'YES' " + _
                                        "WHERE PIPES.DELETED = 0 AND PIPES.FACILITY_ID = " + oInspection.FacilityID.ToString
                Dim ds As DataSet = oInspectionChecklistMasterDB.DBGetDS(strSQL)
                If Not ds Is Nothing Then
                    For Each dr As DataRow In ds.Tables(0).Rows
                        If slPipeFuelType.Contains(dr("PIPE_ID")) Then
                            slPipeFuelType.Item(dr("PIPE_ID")) += ", " + IIf(dr("PROPERTY_NAME") Is DBNull.Value, "N/A", dr("PROPERTY_NAME"))
                        Else
                            slPipeFuelType.Add(dr("PIPE_ID"), IIf(dr("PROPERTY_NAME") Is DBNull.Value, "N/A", dr("PROPERTY_NAME")))
                        End If
                    Next
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Private Sub ResetCLShow()
            Try
                For Each cl As MUSTER.Info.InspectionChecklistMasterInfo In oInspection.ChecklistMasterCollection.Values
                    cl.Show = False
                Next
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Private Sub InitCLItem2point5_6Variables()
            qIDOfCLItem2point5 = 0
            respOfCLItem2point5 = -1
            qIDOfCLItem2point6 = 0
        End Sub
        Private Sub InitCLVariables()
            Try
                hasCIU = False
                hasTOS = False
                hasTOSI = False
                hasPipes = False
                CIUhasPipes = False
                TOShasPipes = False
                TOSIhasPipes = False
                CIUsingleWalledPost88NotInspected = False
                CIU_PPost88NotInspected = False
                CIUfiberGlassPost88NotInspected = False
                CIU_Ppressurized = False
                CIUhazardousSubstance = False
                TOSIhazardousSubstance = False
                CIUUsedOilOnly = False
                CIUballFloat = False
                CIUdropTube = False
                CIUelectronicAlarm = False
                CIUelectronicAlarmOnly = False
                CIUlikeCP = False
                TOSIlikeCP = False
                CIU_PlikeCP = False
                TOSI_PlikeCP = False
                CIUimpressedCurrent = False
                TOSIimpressedCurrent = False
                CIU_PimpressedCurrent = False
                TOSI_PimpressedCurrent = False
                CIU_PtermImpressedCurrent = False
                TOSI_PtermImpressedCurrent = False
                CIUcathodicallyProtected = False
                TOSIcathodicallyProtected = False
                CIULikecathodicallyProtected = False
                TOSILikecathodicallyProtected = False
                CIUlinedInteriorInstallgt10yrsAgo = False
                TOSIlinedInteriorInstallgt10yrsAgo = False
                CIU_PcathodicallyProtected = False
                TOSI_PcathodicallyProtected = False
                CIU_PLikecathodicallyProtected = False
                TOSI_PLikecathodicallyProtected = False
                CIU_Psteel = False
                TOSI_Psteel = False
                CIU_PdispContainedInBoots = False
                TOSI_PdispContainedInBoots = False
                CIU_PtankContainedInBoots = False
                TOSI_PtankContainedInBoots = False
                CIU_PtankContainedInSumps = False
                TOSI_PtankContainedInSumps = False
                CIU_PdispContainedInSumps = False
                TOSI_PdispContainedInSumps = False
                CIU_PdispCathodicallyProtected = False
                CIU_PipeInstalledAfter10_1_08 = False
                CIU_TanksInstalledAfter10_1_08 = False
                TOSI_PdispCathodicallyProtected = False
                CIU_PtankCathodicallyProtected = False
                TOSI_PtankCathodicallyProtected = False
                CIUemergenOnly = False
                CIUgroundWaterVaporMonitoring = False
                CIUinventoryControlPTT = False
                CIUautomaticTankGauging = False
                CIUstatisticalInventoryReconciliation = False
                CIUmanualTankGauging = False
                CIUvisualInterstitialMonitoring = False
                CIUelectronicInterstitialMonitoring = False
                CIU_PpressurizedUSSuction = False
                CIU_PgroundWaterVaporMonitoring = False
                CIU_PLTT = False
                CIU_PUSSuction = False
                CIU_PelectronicALLD = False
                CIU_Pmechanical = False
                CIU_PstatisticalInventoryReconciliation = False
                CIU_PvisualInterstitialMonitoring = False
                CIU_PcontinuousInterstitialMonitoring = False
                CIU_PnotContinuousInterstitialMonitoring = False
                CIU_Pelectronic = False
                CIU_Pplastic = False

                AddCPReadingsTank = False
                AddCPReadingsPipe = False
                AddCPReadingsTerm = False
                CPReadingsTankQID = 0
                CPReadingsPipeQID = 0
                CPReadingsTermQID = 0

                AddMonitorWellsTank = False
                AddMonitorWellsPipe = False
                MonitorWellsTankQID = 0
                MonitorWellsPipeQID = 0
                MonitorWellsTankPipeQID = 0

                nMaxTankCPLineNum = 1
                nMaxPipeCPLineNum = 1
                nMaxTermCPLineNum = 1
                nMaxTankBedMWLineNum = 1
                nMaxLineMWLineNum = 1
                nMaxTankLineMWLineNum = 1
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Private Sub CheckCLVariables()
            Try
                Dim CIUtnkCount As Int64 = 0
                Dim CIUcompCount As Int64 = 0
                Dim TOSItnkCount As Int64 = 0
                Dim TOSIcompCount As Int64 = 0
                Dim CIUtnkpipeCount As Int64 = 0
                Dim TOSItnkpipeCount As Int64 = 0
                Dim epgCount As Int64 = 0
                Dim usedOilCount As Int64 = 0
                Dim ciuSteelCount As Int64 = 0
                Dim tosiSteelCount As Int64 = 0
                Dim tnkInspectedBefore As Boolean = False
                Dim pipeInspectedBefore As Boolean = False
                Dim electronicAlarmCount As Int64 = 0
                Dim insp As BusinessLogic.pInspection
                Dim facLastInspDate As Date = New Date(1900, 1, 1)

                insp = New pInspection
                insp.Retrieve(oInspection.ID, oInspection.StaffID, oInspection.FacilityID, oInspection.OwnerID)
                facLastInspDate = insp.HasfacilityBeenInspected(oInspection.FacilityID)

                If facLastInspDate < New Date(1940, 1, 1) Then
                    facLastInspDate = New Date(1900, 1, 1)
                End If

                For Each tnk As MUSTER.Info.TankInfo In oOwner.Facility.TankCollection.Values

                    tnkInspectedBefore = IIf(Date.Compare(tnk.DateInstalledTank, facLastInspDate) <= 0, True, False)

                    If tnk.TankStatus = 424 Then 'CIU
                        CIUtnkCount += 1
                        hasCIU = True
                        CIUhasPipes = IIf(tnk.pipesCollection.Count > 0, True, hasPipes)
                        CIUsingleWalledPost88NotInspected = IIf((tnk.TankModDesc <> 413 And tnk.TankModDesc <> 415) And (Date.Compare(tnk.DateInstalledTank, CDate("12/22/1988")) > 0 And (Not tnkInspectedBefore)), True, CIUsingleWalledPost88NotInspected)
                        CIUfiberGlassPost88NotInspected = IIf(tnk.TankMatDesc = 348 And Date.Compare(tnk.DateInstalledTank, CDate("12/22/1988")) > 0 And (Not tnkInspectedBefore), True, CIUfiberGlassPost88NotInspected)
                        For Each comp As MUSTER.Info.CompartmentInfo In tnk.CompartmentCollection.Values
                            CIUcompCount += 1
                            'CIUhazardousSubstance = IIf(tnk.Substance = 310, True, CIUhazardousSubstance)
                            'usedOilCount = IIf(tnk.Substance = 314, usedOilCount + 1, usedOilCount)
                            If comp.Substance = 310 Then
                                CIUhazardousSubstance = True
                            ElseIf comp.Substance = 314 Then
                                usedOilCount += 1
                            End If
                        Next
                        CIU_TanksInstalledAfter10_1_08 = IIf(tnk.DateInstalledTank >= New Date(2008, 10, 1), True, CIU_TanksInstalledAfter10_1_08)
                        CIUballFloat = IIf(tnk.OverFillType = 420, True, CIUballFloat)
                        CIUdropTube = IIf(tnk.OverFillType = 419, True, CIUdropTube)
                        electronicAlarmCount = IIf(tnk.OverFillType = 421, electronicAlarmCount + 1, electronicAlarmCount)
                        CIUelectronicAlarm = IIf(tnk.OverFillType = 421, True, CIUelectronicAlarm)
                        CIUlikeCP = IIf(tnk.TankModDesc = 412 Or tnk.TankModDesc = 415 Or tnk.TankModDesc = 475, True, CIUlikeCP)
                        CIUlikeGalvanic = IIf(tnk.TankCPType = 417, True, CIUlikeGalvanic)
                        CIUimpressedCurrent = IIf(tnk.TankCPType = 418, True, CIUimpressedCurrent)
                        CIUcathodicallyProtected = IIf(tnk.TankModDesc = 412, True, CIUcathodicallyProtected)
                        CIULikecathodicallyProtected = IIf(tnk.TankModDesc = 412 Or tnk.TankModDesc = 415 Or tnk.TankModDesc = 475, True, CIULikecathodicallyProtected)
                        CIUlinedInteriorInstallgt10yrsAgo = IIf(tnk.TankModDesc = 476 And Date.Compare(tnk.LinedInteriorInstallDate, DateAdd(DateInterval.Year, -10, Now.Date)) < 0, True, CIUlinedInteriorInstallgt10yrsAgo)
                        epgCount = IIf(tnk.TankEmergen, epgCount + 1, epgCount)
                        CIUgroundWaterVaporMonitoring = IIf(tnk.TankLD = 335, True, CIUgroundWaterVaporMonitoring)
                        CIUinventoryControlPTT = IIf(tnk.TankLD = 338, True, CIUinventoryControlPTT)
                        CIUautomaticTankGauging = IIf(tnk.TankLD = 336, True, CIUautomaticTankGauging)
                        CIUstatisticalInventoryReconciliation = IIf(tnk.TankLD = 340, True, CIUstatisticalInventoryReconciliation)
                        CIUmanualTankGauging = IIf(tnk.TankLD = 337, True, CIUmanualTankGauging)
                        CIUvisualInterstitialMonitoring = IIf(tnk.TankLD = 343, True, CIUvisualInterstitialMonitoring)
                        CIUelectronicInterstitialMonitoring = IIf(tnk.TankLD = 339, True, CIUelectronicInterstitialMonitoring)

                        'For Each pipe As MUSTER.Info.PipeInfo In tnk.pipesCollection.Values
                        '    If pipe.PipeStatusDesc <> 426 Then
                        '        CIUtnkpipeCount += 1
                        '        ciuSteelCount = IIf(pipe.PipeMatDesc = 250 Or pipe.PipeMatDesc = 251, ciuSteelCount + 1, ciuSteelCount)

                        '        pipeInspectedBefore = IIf(Date.Compare(pipe.PipeInstallDate, facLastInspDate) > 0, True, pipeInspectedBefore)

                        '        CIU_PPost88NotInspected = IIf(Date.Compare(pipe.PipeInstallDate, CDate("12/22/1988")) > 0 And (Not pipeInspectedBefore), True, CIU_PPost88NotInspected)
                        '        CIU_Ppressurized = IIf(pipe.PipeTypeDesc = 266, True, CIU_Ppressurized)
                        '        CIU_PlikeCP = IIf(pipe.PipeModDesc = 260 Or pipe.PipeModDesc = 263, True, CIU_PlikeCP)
                        '        CIU_PimpressedCurrent = IIf(pipe.PipeCPType = 478, True, CIU_PimpressedCurrent)
                        '        CIU_PtermImpressedCurrent = IIf(pipe.TermCPTypeDisp = 478 Or pipe.TermCPTypeTank = 478, True, CIU_PtermImpressedCurrent)
                        '        CIU_PcathodicallyProtected = IIf(pipe.PipeModDesc = 260, True, CIU_PcathodicallyProtected)
                        '        CIU_PdispContainedInBoots = IIf(pipe.TermTypeDisp = 490, True, CIU_PdispContainedInBoots)
                        '        CIU_PtankContainedInBoots = IIf(pipe.TermTypeTank = 483, True, CIU_PtankContainedInBoots)
                        '        'CIU_PdispContainedInSumps = IIf(pipe.TermTypeDisp = 489, True, CIU_PdispContainedInSumps)
                        '        'CIU_PtankContainedInSumps = IIf(pipe.TermTypeTank = 482, True, CIU_PtankContainedInSumps)
                        '        CIU_PdispContainedInSumps = IIf(pipe.ContainSumpDisp, True, CIU_PdispContainedInSumps)
                        '        CIU_PtankContainedInSumps = IIf(pipe.ContainSumpTank, True, CIU_PtankContainedInSumps)
                        '        CIU_PdispCathodicallyProtected = IIf(pipe.TermTypeDisp = 611 Or pipe.TermTypeDisp = 488, True, CIU_PdispCathodicallyProtected)
                        '        CIU_PtankCathodicallyProtected = IIf(pipe.TermTypeTank = 610 Or pipe.TermTypeTank = 481, True, CIU_PtankCathodicallyProtected)
                        '        CIU_PpressurizedUSSuction = IIf(pipe.PipeTypeDesc = 266 Or pipe.PipeTypeDesc = 268, True, CIU_PpressurizedUSSuction)
                        '        CIU_PgroundWaterVaporMonitoring = IIf(pipe.PipeLD = 241, True, CIU_PgroundWaterVaporMonitoring)
                        '        CIU_PLTT = IIf(pipe.PipeLD = 245, True, CIU_PLTT)
                        '        CIU_PUSSuction = IIf(pipe.PipeTypeDesc = 268, True, CIU_PUSSuction)
                        '        CIU_PelectronicALLD = IIf(pipe.PipeLD = 246, True, CIU_PelectronicALLD)
                        '        CIU_Pmechanical = IIf(pipe.ALLDType = 496, True, CIU_Pmechanical)
                        '        CIU_PstatisticalInventoryReconciliation = IIf(pipe.PipeLD = 244, True, CIU_PstatisticalInventoryReconciliation)
                        '        CIU_PvisualInterstitialMonitoring = IIf(pipe.PipeLD = 242, True, CIU_PvisualInterstitialMonitoring)
                        '        CIU_PcontinuousInterstitialMonitoring = IIf(pipe.PipeLD = 243, True, CIU_PcontinuousInterstitialMonitoring)
                        '        CIU_PnotContinuousInterstitialMonitoring = IIf(pipe.PipeLD <> 243, True, CIU_PnotContinuousInterstitialMonitoring)
                        '        CIU_Pelectronic = IIf(pipe.ALLDType = 497, True, CIU_Pelectronic)
                        '        CIU_Pplastic = IIf(pipe.PipeMatDesc = 252 Or pipe.PipeMatDesc = 254 Or pipe.PipeMatDesc = 494, True, CIU_Pplastic)
                        '    End If
                        'Next
                    ElseIf tnk.TankStatus = 425 Then ' TOS
                        hasTOS = True
                        TOShasPipes = IIf(tnk.pipesCollection.Count > 0, True, TOShasPipes)
                    ElseIf tnk.TankStatus = 429 Then ' TOSI
                        TOSItnkCount += 1
                        hasTOSI = True
                        TOSIhasPipes = IIf(tnk.pipesCollection.Count > 0, True, TOSIhasPipes)
                        For Each comp As MUSTER.Info.CompartmentInfo In tnk.CompartmentCollection.Values
                            TOSIcompCount += 1
                            'TOSIhazardousSubstance = IIf(tnk.Substance = 310, True, TOSIhazardousSubstance)
                            If comp.Substance = 310 Then
                                TOSIhazardousSubstance = True
                                Exit For
                            End If
                        Next
                        TOSIlikeCP = IIf(tnk.TankModDesc = 412 Or tnk.TankModDesc = 415 Or tnk.TankModDesc = 475, True, TOSIlikeCP)
                        Me.TOSlikeGalvanic = IIf(tnk.TankModDesc = 417, True, TOSlikeGalvanic)

                        TOSIimpressedCurrent = IIf(tnk.TankCPType = 418, True, TOSIimpressedCurrent)
                        TOSIcathodicallyProtected = IIf(tnk.TankModDesc = 412, True, TOSIcathodicallyProtected)
                        TOSILikecathodicallyProtected = IIf(tnk.TankModDesc = 412 Or tnk.TankModDesc = 415 Or tnk.TankModDesc = 475, True, TOSILikecathodicallyProtected)
                        TOSIlinedInteriorInstallgt10yrsAgo = IIf(tnk.TankModDesc = 476 And Date.Compare(tnk.LinedInteriorInstallDate, DateAdd(DateInterval.Year, -10, Now.Date)) < 0, True, TOSIlinedInteriorInstallgt10yrsAgo)

                        'For Each pipe As MUSTER.Info.PipeInfo In tnk.pipesCollection.Values
                        '    TOSItnkpipeCount += 1
                        '    tosiSteelCount = IIf(pipe.PipeMatDesc = 250 Or pipe.PipeMatDesc = 251, ciuSteelCount + 1, ciuSteelCount)

                        '    TOSI_PlikeCP = IIf(pipe.PipeModDesc = 260 Or pipe.PipeModDesc = 263, True, TOSI_PlikeCP)
                        '    TOSI_PimpressedCurrent = IIf(pipe.PipeCPType = 478, True, TOSI_PimpressedCurrent)
                        '    TOSI_PtermImpressedCurrent = IIf(pipe.TermCPTypeDisp = 478 Or pipe.TermCPTypeTank = 478, True, TOSI_PtermImpressedCurrent)
                        '    TOSI_PcathodicallyProtected = IIf(pipe.PipeModDesc = 260, True, TOSI_PcathodicallyProtected)
                        '    TOSI_PdispContainedInBoots = IIf(pipe.TermTypeDisp = 490, True, TOSI_PdispContainedInBoots)
                        '    TOSI_PtankContainedInBoots = IIf(pipe.TermTypeTank = 483, True, TOSI_PtankContainedInBoots)
                        '    'TOSI_PdispContainedInSumps = IIf(pipe.TermTypeDisp = 489, True, TOSI_PdispContainedInSumps)
                        '    'TOSI_PtankContainedInSumps = IIf(pipe.TermTypeTank = 482, True, TOSI_PtankContainedInSumps)
                        '    TOSI_PdispContainedInSumps = IIf(pipe.ContainSumpDisp, True, TOSI_PdispContainedInSumps)
                        '    TOSI_PtankContainedInSumps = IIf(pipe.ContainSumpTank, True, TOSI_PtankContainedInSumps)
                        '    TOSI_PdispCathodicallyProtected = IIf(pipe.TermTypeDisp = 611 Or pipe.TermTypeDisp = 488, True, TOSI_PdispCathodicallyProtected)
                        '    TOSI_PtankCathodicallyProtected = IIf(pipe.TermTypeTank = 610 Or pipe.TermTypeTank = 481, True, TOSI_PtankCathodicallyProtected)
                        'Next
                    End If

                    For Each pipe As MUSTER.Info.PipeInfo In tnk.pipesCollection.Values
                        If pipe.PipeStatusDesc <> 426 And pipe.PipeStatusDesc = 424 Then  'CIU
                            CIUtnkpipeCount += 1
                            ciuSteelCount = IIf(pipe.PipeMatDesc = 250 Or pipe.PipeMatDesc = 251, ciuSteelCount + 1, ciuSteelCount)

                            pipeInspectedBefore = IIf(Date.Compare(pipe.PipeInstallDate, facLastInspDate) <= 0, True, pipeInspectedBefore)

                            CIU_PPost88NotInspected = IIf(Date.Compare(pipe.PipeInstallDate, CDate("12/22/1988")) > 0 And (Not pipeInspectedBefore), True, CIU_PPost88NotInspected)

                            CIU_PipeInstalledAfter10_1_08 = IIf(pipe.PipeInstallDate >= New Date(2008, 10, 1), True, CIU_PipeInstalledAfter10_1_08)

                            CIU_Ppressurized = IIf(pipe.PipeTypeDesc = 266, True, CIU_Ppressurized)
                            CIU_PlikeCP = IIf(pipe.PipeModDesc = 260 Or pipe.PipeModDesc = 263, True, CIU_PlikeCP)
                            CIU_PimpressedCurrent = IIf(pipe.PipeCPType = 478, True, CIU_PimpressedCurrent)
                            CIU_PtermImpressedCurrent = IIf(pipe.TermCPTypeDisp = 478 Or pipe.TermCPTypeTank = 478, True, CIU_PtermImpressedCurrent)
                            CIU_PcathodicallyProtected = IIf(pipe.PipeModDesc = 260, True, CIU_PcathodicallyProtected)
                            CIU_PLikecathodicallyProtected = IIf(pipe.PipeModDesc = 260 Or pipe.PipeModDesc = 263, True, CIU_PLikecathodicallyProtected)
                            CIU_PdispContainedInBoots = IIf(pipe.TermTypeDisp = 490, True, CIU_PdispContainedInBoots)
                            CIU_PtankContainedInBoots = IIf(pipe.TermTypeTank = 483, True, CIU_PtankContainedInBoots)
                            CIU_PdispContainedInSumps = IIf(pipe.TermTypeDisp = 489, True, CIU_PdispContainedInSumps)
                            CIU_PtankContainedInSumps = IIf(pipe.TermTypeTank = 482, True, CIU_PtankContainedInSumps)
                            CIU_PdispContainedInSumps = IIf(pipe.ContainSumpDisp, True, CIU_PdispContainedInSumps)
                            CIU_PtankContainedInSumps = IIf(pipe.ContainSumpTank, True, CIU_PtankContainedInSumps)
                            CIU_PdispCathodicallyProtected = IIf(pipe.TermTypeDisp = 611 Or pipe.TermTypeDisp = 488, True, CIU_PdispCathodicallyProtected)
                            CIU_PtankCathodicallyProtected = IIf(pipe.TermTypeTank = 610 Or pipe.TermTypeTank = 481, True, CIU_PtankCathodicallyProtected)
                            CIU_PpressurizedUSSuction = IIf(pipe.PipeTypeDesc = 266 Or pipe.PipeTypeDesc = 268, True, CIU_PpressurizedUSSuction)
                            CIU_PgroundWaterVaporMonitoring = IIf(pipe.PipeLD = 241, True, CIU_PgroundWaterVaporMonitoring)
                            CIU_PLTT = IIf(pipe.PipeLD = 245, True, CIU_PLTT)
                            CIU_PUSSuction = IIf(pipe.PipeTypeDesc = 268, True, CIU_PUSSuction)
                            CIU_PelectronicALLD = IIf(pipe.PipeLD = 246, True, CIU_PelectronicALLD)
                            CIU_Pmechanical = IIf(pipe.ALLDType = 496, True, CIU_Pmechanical)
                            CIU_PstatisticalInventoryReconciliation = IIf(pipe.PipeLD = 244, True, CIU_PstatisticalInventoryReconciliation)
                            CIU_PvisualInterstitialMonitoring = IIf(pipe.PipeLD = 242, True, CIU_PvisualInterstitialMonitoring)
                            CIU_PcontinuousInterstitialMonitoring = IIf(pipe.PipeLD = 243, True, CIU_PcontinuousInterstitialMonitoring)
                            CIU_PnotContinuousInterstitialMonitoring = IIf(pipe.PipeLD <> 243, True, CIU_PnotContinuousInterstitialMonitoring)
                            CIU_Pelectronic = IIf(pipe.ALLDType = 497, True, CIU_Pelectronic)
                            CIU_Pplastic = IIf(pipe.PipeMatDesc = 252 Or pipe.PipeMatDesc = 254 Or pipe.PipeMatDesc = 494, True, CIU_Pplastic)
                        ElseIf pipe.PipeStatusDesc <> 426 And pipe.PipeStatusDesc = 429 Then ' TOSI
                            TOSItnkpipeCount += 1
                            tosiSteelCount = IIf(pipe.PipeMatDesc = 250 Or pipe.PipeMatDesc = 251, ciuSteelCount + 1, ciuSteelCount)

                            TOSI_PlikeCP = IIf(pipe.PipeModDesc = 260 Or pipe.PipeModDesc = 263, True, TOSI_PlikeCP)
                            TOSI_PimpressedCurrent = IIf(pipe.PipeCPType = 478, True, TOSI_PimpressedCurrent)
                            TOSI_PtermImpressedCurrent = IIf(pipe.TermCPTypeDisp = 478 Or pipe.TermCPTypeTank = 478, True, TOSI_PtermImpressedCurrent)
                            TOSI_PcathodicallyProtected = IIf(pipe.PipeModDesc = 260, True, TOSI_PcathodicallyProtected)
                            TOSI_PLikecathodicallyProtected = IIf(pipe.PipeModDesc = 260 Or pipe.PipeModDesc = 263, True, TOSI_PLikecathodicallyProtected)
                            TOSI_PdispContainedInBoots = IIf(pipe.TermTypeDisp = 490, True, TOSI_PdispContainedInBoots)
                            TOSI_PtankContainedInBoots = IIf(pipe.TermTypeTank = 483, True, TOSI_PtankContainedInBoots)
                            TOSI_PdispContainedInSumps = IIf(pipe.TermTypeDisp = 489, True, TOSI_PdispContainedInSumps)
                            TOSI_PtankContainedInSumps = IIf(pipe.TermTypeTank = 482, True, TOSI_PtankContainedInSumps)
                            TOSI_PdispContainedInSumps = IIf(pipe.ContainSumpDisp, True, TOSI_PdispContainedInSumps)
                            TOSI_PtankContainedInSumps = IIf(pipe.ContainSumpTank, True, TOSI_PtankContainedInSumps)
                            TOSI_PdispCathodicallyProtected = IIf(pipe.TermTypeDisp = 611 Or pipe.TermTypeDisp = 488, True, TOSI_PdispCathodicallyProtected)
                            TOSI_PtankCathodicallyProtected = IIf(pipe.TermTypeTank = 610 Or pipe.TermTypeTank = 481, True, TOSI_PtankCathodicallyProtected)
                        End If
                    Next

                Next

                CIUUsedOilOnly = IIf(usedOilCount = CIUcompCount And CIUcompCount <> 0, True, CIUUsedOilOnly)
                CIUemergenOnly = IIf(epgCount = CIUtnkCount And CIUtnkCount <> 0, True, CIUemergenOnly)
                CIU_Psteel = IIf(ciuSteelCount = CIUtnkpipeCount And CIUtnkpipeCount <> 0, True, CIU_Psteel)
                TOSI_Psteel = IIf(tosiSteelCount = TOSItnkpipeCount And TOSItnkpipeCount <> 0, True, TOSI_Psteel)
                CIUelectronicAlarmOnly = IIf(electronicAlarmCount = CIUtnkCount And CIUtnkCount <> 0, True, CIUelectronicAlarmOnly)

                ' if there are no cp readings in the collection, it was never created
                ' hence the line numbering starts with 1
                ' else get the max number from db

                nMaxTankCPLineNum = 1
                nMaxPipeCPLineNum = 1
                nMaxTermCPLineNum = 1

                If oInspection.CPReadingsCollection.Count > 0 Then
                    Dim ds As DataSet
                    Dim strSQL As String = ""
                    ' need to check deleted records cause line num is not reused
                    ' max tank cp line num
                    strSQL = "SELECT ISNULL(MAX(LINE_NUMBER),0) AS MAX_LINE_NUM FROM tblINS_INSPECTION_CP_READINGS WHERE INSPECTION_ID = " + oInspection.ID.ToString + " AND QUESTION_ID = 31"
                    ds = oInspectionChecklistMasterDB.DBGetDS(strSQL)
                    If ds.Tables(0).Rows.Count > 0 Then
                        nMaxTankCPLineNum = ds.Tables(0).Rows(0)("MAX_LINE_NUM")
                        'If nMaxTankCPLineNum < 0 Then nMaxTankCPLineNum = 0
                    End If

                    ' max pipe cp line num
                    strSQL = "SELECT ISNULL(MAX(LINE_NUMBER),0) AS MAX_LINE_NUM FROM tblINS_INSPECTION_CP_READINGS WHERE INSPECTION_ID = " + oInspection.ID.ToString + " AND QUESTION_ID = 36"
                    ds = oInspectionChecklistMasterDB.DBGetDS(strSQL)
                    If ds.Tables(0).Rows.Count > 0 Then
                        nMaxPipeCPLineNum = ds.Tables(0).Rows(0)("MAX_LINE_NUM")
                        'If nMaxPipeCPLineNum < 0 Then nMaxPipeCPLineNum = 0
                    End If

                    ' max term cp line num
                    strSQL = "SELECT ISNULL(MAX(LINE_NUMBER),0) AS MAX_LINE_NUM FROM tblINS_INSPECTION_CP_READINGS WHERE INSPECTION_ID = " + oInspection.ID.ToString + " AND QUESTION_ID = 44"
                    ds = oInspectionChecklistMasterDB.DBGetDS(strSQL)
                    If ds.Tables(0).Rows.Count > 0 Then
                        nMaxTermCPLineNum = ds.Tables(0).Rows(0)("MAX_LINE_NUM")
                        'If nMaxTermCPLineNum < 0 Then nMaxTermCPLineNum = 0
                    End If

                    ' looping through the collection to get the max
                    For Each cp As MUSTER.Info.InspectionCPReadingsInfo In oInspection.CPReadingsCollection.Values
                        If cp.QuestionID = CPReadingsTankQID Then
                            If Not cp.RemoteReferCellPlacement And Not cp.GalvanicIC Then
                                If cp.LineNumber > nMaxTankCPLineNum Then
                                    nMaxTankCPLineNum = cp.LineNumber
                                End If
                            End If
                        ElseIf cp.QuestionID = CPReadingsPipeQID Then
                            If Not cp.RemoteReferCellPlacement And Not cp.GalvanicIC Then
                                If cp.LineNumber > nMaxTankCPLineNum Then
                                    nMaxPipeCPLineNum = cp.LineNumber
                                End If
                            End If
                        ElseIf cp.QuestionID = CPReadingsTermQID Then
                            If Not cp.RemoteReferCellPlacement And Not cp.GalvanicIC Then
                                If cp.LineNumber > nMaxTankCPLineNum Then
                                    nMaxTermCPLineNum = cp.LineNumber
                                End If
                            End If
                        End If
                    Next
                    nMaxTankCPLineNum += 1
                    nMaxPipeCPLineNum += 1
                    nMaxTermCPLineNum += 1
                End If

                ' if there are no mw in the collection, it was never created
                ' hence the line numbering starts with 1
                ' else get the max number from db

                nMaxTankBedMWLineNum = 0
                nMaxLineMWLineNum = 0
                nMaxTankLineMWLineNum = 0

                If oInspection.MonitorWellsCollection.Count > 0 Then
                    Dim ds As DataSet
                    Dim strSQL As String = ""
                    ' need to check deleted records cause line num is not reused
                    ' max tankbed mw line num
                    strSQL = "SELECT ISNULL(MAX(LINE_NUMBER),0) AS MAX_LINE_NUM FROM tblINS_INSPECTION_MONITOR_WELLS WHERE INSPECTION_ID = " + oInspection.ID.ToString + " AND QUESTION_ID = 56"
                    ds = oInspectionChecklistMasterDB.DBGetDS(strSQL)
                    If ds.Tables(0).Rows.Count > 0 Then
                        nMaxTankBedMWLineNum = ds.Tables(0).Rows(0)("MAX_LINE_NUM")
                    End If

                    ' max line mw line num
                    strSQL = "SELECT ISNULL(MAX(LINE_NUMBER),0) AS MAX_LINE_NUM FROM tblINS_INSPECTION_MONITOR_WELLS WHERE INSPECTION_ID = " + oInspection.ID.ToString + " AND QUESTION_ID = 94"
                    ds = oInspectionChecklistMasterDB.DBGetDS(strSQL)
                    If ds.Tables(0).Rows.Count > 0 Then
                        nMaxLineMWLineNum = ds.Tables(0).Rows(0)("MAX_LINE_NUM")
                    End If

                    ' max tank line mw line num
                    strSQL = "SELECT ISNULL(MAX(LINE_NUMBER),0) AS MAX_LINE_NUM FROM tblINS_INSPECTION_MONITOR_WELLS WHERE INSPECTION_ID = " + oInspection.ID.ToString + " AND QUESTION_ID = 141"
                    ds = oInspectionChecklistMasterDB.DBGetDS(strSQL)
                    If ds.Tables(0).Rows.Count > 0 Then
                        nMaxTankLineMWLineNum = ds.Tables(0).Rows(0)("MAX_LINE_NUM")
                    End If

                    ' looping through the collection to get the max
                    For Each mwell As MUSTER.Info.InspectionMonitorWellsInfo In oInspection.MonitorWellsCollection.Values
                        If mwell.TankLine Then
                            If mwell.LineNumber > nMaxTankBedMWLineNum Then
                                nMaxTankBedMWLineNum = mwell.LineNumber
                            End If
                        ElseIf mwell.QuestionID = 141 Then
                            If mwell.LineNumber > nMaxTankLineMWLineNum Then
                                nMaxTankLineMWLineNum = mwell.LineNumber
                            End If
                        Else
                            If mwell.LineNumber > nMaxLineMWLineNum Then
                                nMaxLineMWLineNum = mwell.LineNumber
                            End If
                        End If
                    Next
                End If
                nMaxTankBedMWLineNum += 1
                nMaxLineMWLineNum += 1
                nMaxTankLineMWLineNum += 1
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Private Sub SetInspectionToChild()
            Try
                oInspectionResponses.InspectionInfo = oInspection
                oInspectionCPReadings.InspectionInfo = oInspection
                oInspectionMonitorWells.InspectionInfo = oInspection
                oInspectionCCAT.InspectionInfo = oInspection
                oInspectionCitation.InspectionInfo = oInspection
                oInspectionDiscrep.InspectionInfo = oInspection
                oInspectionRectifier.InspectionInfo = oInspection
                oInspectionSketch.InspectionInfo = oInspection
                oInspectionSOC.InspectionInfo = oInspection
                oInspectionComments.InspectionInfo = oInspection
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Private Sub CheckTankPipeBelongToCL(Optional ByVal tankInfo As MUSTER.Info.TankInfo = Nothing, Optional ByVal pipeInfo As MUSTER.Info.PipeInfo = Nothing, Optional ByVal [readOnly] As Boolean = False)
            ' Sub to assign tank/pipe id which meet the conditions to respective checklist item
            ' check if checklist is visible
            ' if visible, check tank / pipes for condition
            Try
                If Not [readOnly] Then
                    If tankInfo Is Nothing Then
                        For Each clInfo As MUSTER.Info.InspectionChecklistMasterInfo In oInspection.ChecklistMasterCollection.Values
                            clInfo.TankArrayList.Clear()
                            clInfo.PipeArrayList.Clear()
                            clInfo.PipeTermArrayList.Clear()
                        Next
                        For Each tnk As MUSTER.Info.TankInfo In oOwner.Facility.TankCollection.Values
                            If Not htTankIDIndex.Contains(tnk.TankId) Then
                                htTankIDIndex.Add(tnk.TankId, tnk.TankIndex)
                            End If
                            If tnk.TankStatus <> 426 Then
                                SetTankPipeBelongToCCAT(tnk, , [readOnly])
                            End If
                            For Each pipe As MUSTER.Info.PipeInfo In tnk.pipesCollection.Values
                                SetTankPipeBelongToCCAT(tnk, pipe, [readOnly])
                                If Not htPipeIDIndex.Contains(pipe.PipeID) Then
                                    htPipeIDIndex.Add(pipe.PipeID, pipe.Index)
                                End If
                            Next
                        Next
                    ElseIf pipeInfo Is Nothing Then
                        SetTankPipeBelongToCCAT(tankInfo, , [readOnly])
                    Else
                        SetTankPipeBelongToCCAT(tankInfo, pipeInfo, [readOnly])
                    End If
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Private Sub SetTankPipeBelongToCCAT(ByVal tnk As MUSTER.Info.TankInfo, Optional ByVal pipe As MUSTER.Info.PipeInfo = Nothing, Optional ByVal [readOnly] As Boolean = False)
            Try
                If Not [readOnly] Then
                    Dim facLastInspDate As Date = IIf(Date.Compare(oInspection.RescheduledDate, CDate("01/01/0001")) = 0, oInspection.ScheduledDate, oInspection.RescheduledDate)
                    Dim tnkInspectedBefore As Boolean = IIf(Date.Compare(tnk.DateInstalledTank, facLastInspDate) > 0, True, False)
                    Dim pipeInspectedBefore As Boolean = False
                    If Not (pipe Is Nothing) Then
                        pipeInspectedBefore = IIf(Date.Compare(pipe.PipeInstallDate, facLastInspDate) > 0, True, False)
                    End If

                    For Each oInspCLInfo As MUSTER.Info.InspectionChecklistMasterInfo In oInspection.ChecklistMasterCollection.Values
                        If oInspCLInfo.Show And (oInspCLInfo.AppliesToTank Or oInspCLInfo.AppliesToPipe Or oInspCLInfo.AppliesToPipeTerm) Then
                            ' check tank / pipe for condition
                            Select Case oInspCLInfo.CheckListItemNumber.Trim
                                Case "1.3"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 Or tnk.TankStatus = 429 Or tnk.TankStatus = 425 Then 'CIU OR TOSI OR TOS
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "1.4"
                                    If Not (pipe Is Nothing) Then
                                        'If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                        '    oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                        '    'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                        '    '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                        '    'End If
                                        'End If
                                        If oInspCLInfo.AppliesToPipe Then
                                            If (tnk.TankStatus = 424 Or tnk.TankStatus = 429 Or tnk.TankStatus = 425) And tnk.pipesCollection.Count > 0 Then 'CIU OR TOSI OR TOS
                                                If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                    oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                                End If
                                                'If Not colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                                '    colPipeIDIndex.Add(pipe.PipeID.ToString, pipe.Index)
                                                'End If
                                            End If
                                        End If
                                    End If
                                Case "1.5"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And (tnk.TankModDesc <> 413 Or tnk.TankModDesc <> 415) And (Date.Compare(tnk.DateInstalledTank, CDate("12/22/1988")) > 0 And (Not tnkInspectedBefore)) Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "1.6"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And Date.Compare(pipe.PipeInstallDate, CDate("12/22/1988")) > 0 And (Not pipeInspectedBefore) Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                                'If Not colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                                '    colPipeIDIndex.Add(pipe.PipeID.ToString, pipe.Index)
                                                'End If
                                            End If
                                        End If
                                    End If
                                Case "1.7"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.TankMatDesc = 348 And Date.Compare(tnk.DateInstalledTank, CDate("12/22/1988")) > 0 And (Not tnkInspectedBefore) Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "1.8"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And pipe.PipeTypeDesc = 266 Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                                'If Not colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                                '    colPipeIDIndex.Add(pipe.PipeID.ToString, pipe.Index)
                                                'End If
                                            End If
                                        End If
                                    End If
                                Case "1.9"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If (tnk.TankStatus = 424 Or tnk.TankStatus = 429 Or tnk.TankStatus = 425) Then
                                            For Each comp As MUSTER.Info.CompartmentInfo In tnk.CompartmentCollection.Values
                                                If comp.Substance = 310 Then
                                                    If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                        oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                                    End If
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                    End If
                                Case "1.10"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And pipe.PipeTypeDesc = 266 Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                                'If Not colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                                '    colPipeIDIndex.Add(pipe.PipeID.ToString, pipe.Index)
                                                'End If
                                            End If
                                        End If
                                    End If
                                Case "1.10.1", "1.10.2"
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And pipe.PipeTypeDesc = 266 Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                                'If Not colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                                '    colPipeIDIndex.Add(pipe.PipeID.ToString, pipe.Index)
                                                'End If
                                            End If
                                        End If
                                    End If

                                Case "1.11"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 AndAlso (Date.Compare(tnk.DateInstalledTank, CDate("9/30/2008")) > 0) Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "1.12"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToPipe Then
                                        If Not (pipe Is Nothing) Then
                                            If pipe.PipeStatusDesc = 424 AndAlso (Date.Compare(pipe.PipeInstallDate, CDate("9/30/2008")) > 0) Then
                                                If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                    oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                                End If
                                            End If
                                        End If

                                    End If
                                Case "1.16", "1.16.1", "1.17", "1.17.1"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToPipe Then
                                        If Not (pipe Is Nothing) Then
                                            If pipe.PipeStatusDesc = 424 AndAlso (Date.Compare(pipe.PipeInstallDate, CDate("9/30/2008")) > 0) Then
                                                If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                    oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                                End If
                                            End If
                                        End If

                                    End If
                                Case "1.15", "1.15.1"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 AndAlso (Date.Compare(tnk.DateInstalledTank, CDate("9/30/2008")) > 0) Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If

                                    End If

                                Case "2.1"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "2.2"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 Then
                                            For Each comp As MUSTER.Info.CompartmentInfo In tnk.CompartmentCollection.Values
                                                If comp.Substance <> 314 Then
                                                    If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                        oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                                    End If
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                    End If
                                Case "2.3"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 Then
                                            For Each comp As MUSTER.Info.CompartmentInfo In tnk.CompartmentCollection.Values
                                                If comp.Substance <> 314 Then
                                                    If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                        oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                                    End If
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                    End If
                                Case "2.3.1"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 Then
                                            For Each comp As MUSTER.Info.CompartmentInfo In tnk.CompartmentCollection.Values
                                                If comp.Substance <> 314 Then
                                                    If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                        oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                                    End If
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                    End If
                                Case "2.4"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 Then
                                            For Each comp As MUSTER.Info.CompartmentInfo In tnk.CompartmentCollection.Values
                                                If comp.Substance <> 314 Then
                                                    If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                        oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                                    End If
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                    End If
                                Case "2.5"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.OverFillType = 420 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "2.6"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.OverFillType = 420 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "2.7"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.OverFillType = 419 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "2.8"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.OverFillType = 421 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "2.9"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.OverFillType = 420 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "2.10"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 Then
                                            Dim bolAddTankID As Boolean = False
                                            For Each comp As MUSTER.Info.CompartmentInfo In tnk.CompartmentCollection.Values
                                                If comp.Substance <> 314 Then
                                                    bolAddTankID = True
                                                    Exit For
                                                End If
                                            Next
                                            If tnk.OverFillType <> 421 Or bolAddTankID Then
                                                If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                    oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                                End If
                                            End If
                                        End If
                                    End If
                                Case "2.11", "2.12", "2.11.1", "2.12.1"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 Then
                                            For Each comp As MUSTER.Info.CompartmentInfo In tnk.CompartmentCollection.Values
                                                If comp.Substance <> 314 Then
                                                    If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                        oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                                    End If
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                    End If
                                Case "2.13", "2.14"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 Then
                                            For Each comp As MUSTER.Info.CompartmentInfo In tnk.CompartmentCollection.Values
                                                If comp.Substance <> 314 Then
                                                    If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                        oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                                    End If
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                    End If
                                Case "2.15"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 Then
                                            For Each comp As MUSTER.Info.CompartmentInfo In tnk.CompartmentCollection.Values
                                                If comp.Substance <> 314 Then
                                                    If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                        oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                                    End If
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                    End If
                                Case "2.16"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.OverFillType = 420 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "3.1"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    '    If oInspCLInfo.PipeTermArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeTermArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToTank And oInspCLInfo.AppliesToPipe And oInspCLInfo.AppliesToPipeTerm Then
                                        If (tnk.TankStatus = 424 Or tnk.TankStatus = 429 Or tnk.TankStatus = 425) And (tnk.TankModDesc = 412 Or tnk.TankModDesc = 415 Or tnk.TankModDesc = 475) Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                        If Not (pipe Is Nothing) Then
                                            If (tnk.TankStatus = 424 Or tnk.TankStatus = 429 Or tnk.TankStatus = 425) And (pipe.PipeModDesc = 260 Or pipe.PipeModDesc = 263) Then
                                                If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                    oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                                End If

                                            End If

                                            If (tnk.TankStatus = 424 Or tnk.TankStatus = 429 Or tnk.TankStatus = 425) And (pipe.TermTypeDisp = 611 Or pipe.TermTypeTank = 610) Then
                                                If Not oInspCLInfo.PipeTermArrayList.Contains(pipe.PipeID) Then
                                                    oInspCLInfo.PipeTermArrayList.Add(pipe.PipeID)
                                                End If
                                            End If
                                        End If

                                    End If

                                Case "3.2"
                                    If oInspCLInfo.AppliesToTank And oInspCLInfo.AppliesToPipe And oInspCLInfo.AppliesToPipeTerm Then
                                        If (tnk.TankStatus = 424 Or tnk.TankStatus = 429 Or tnk.TankStatus = 425) And (tnk.TankModDesc = 412 Or tnk.TankModDesc = 415 Or tnk.TankModDesc = 475) Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                        If Not (pipe Is Nothing) Then
                                            If (tnk.TankStatus = 424 Or tnk.TankStatus = 429 Or tnk.TankStatus = 425) And (pipe.PipeModDesc = 260 Or pipe.PipeModDesc = 263) Then
                                                If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                    oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                                End If

                                            End If

                                            If (tnk.TankStatus = 424 Or tnk.TankStatus = 429 Or tnk.TankStatus = 425) And (pipe.TermTypeDisp = 611 Or pipe.TermTypeTank = 610) Then
                                                If Not oInspCLInfo.PipeTermArrayList.Contains(pipe.PipeID) Then
                                                    oInspCLInfo.PipeTermArrayList.Add(pipe.PipeID)
                                                End If
                                            End If
                                        End If
                                    End If
                                Case "3.2.1"
                                    If oInspCLInfo.AppliesToTank Then
                                        If (tnk.TankStatus = 424 Or tnk.TankStatus = 429 Or tnk.TankStatus = 425) And (tnk.TankCPType = 417) Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If

                                Case "3.3"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    '    If oInspCLInfo.PipeTermArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeTermArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToTank And oInspCLInfo.AppliesToPipe And oInspCLInfo.AppliesToPipeTerm Then
                                        If (tnk.TankStatus = 424 Or tnk.TankStatus = 429 Or tnk.TankStatus = 425) And (tnk.TankModDesc = 412 Or tnk.TankModDesc = 415 Or tnk.TankModDesc = 475) And (tnk.TankCPType = 418) Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                        If Not (pipe Is Nothing) Then
                                            If (tnk.TankStatus = 424 Or tnk.TankStatus = 429 Or tnk.TankStatus = 425) And (pipe.PipeModDesc = 260 Or pipe.PipeModDesc = 263) AndAlso pipe.PipeCPType = 478 Then
                                                If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                    oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                                End If

                                            End If

                                            If (tnk.TankStatus = 424 Or tnk.TankStatus = 429 Or tnk.TankStatus = 425) And (pipe.TermTypeDisp = 611 Or pipe.TermTypeTank = 610) AndAlso (pipe.TermCPTypeDisp = 478 Or pipe.TermCPTypeTank = 478) Then
                                                If Not oInspCLInfo.PipeTermArrayList.Contains(pipe.PipeID) Then
                                                    oInspCLInfo.PipeTermArrayList.Add(pipe.PipeID)
                                                End If
                                            End If
                                        End If
                                    End If
                                Case "3.4"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    '    If oInspCLInfo.PipeTermArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeTermArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToTank And oInspCLInfo.AppliesToPipe And oInspCLInfo.AppliesToPipeTerm Then
                                        If (tnk.TankStatus = 424 Or tnk.TankStatus = 429 Or tnk.TankStatus = 425) And (tnk.TankModDesc = 412 Or tnk.TankModDesc = 415 Or tnk.TankModDesc = 475) And (tnk.TankCPType = 418) Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                        If Not (pipe Is Nothing) Then
                                            If (tnk.TankStatus = 424 Or tnk.TankStatus = 429 Or tnk.TankStatus = 425) And (pipe.PipeModDesc = 260 Or pipe.PipeModDesc = 263) AndAlso pipe.PipeCPType = 478 Then
                                                If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                    oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                                End If

                                            End If

                                            If (tnk.TankStatus = 424 Or tnk.TankStatus = 429 Or tnk.TankStatus = 425) And (pipe.TermTypeDisp = 611 Or pipe.TermTypeTank = 610) AndAlso (pipe.TermCPTypeDisp = 478 Or pipe.TermCPTypeTank = 478) Then
                                                If Not oInspCLInfo.PipeTermArrayList.Contains(pipe.PipeID) Then
                                                    oInspCLInfo.PipeTermArrayList.Add(pipe.PipeID)
                                                End If
                                            End If
                                        End If
                                    End If
                                Case "3.5.1"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 Or tnk.TankStatus = 429 Or tnk.TankStatus = 425 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "3.5.2"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If (tnk.TankStatus = 424 Or tnk.TankStatus = 429 Or tnk.TankStatus = 425) And (tnk.TankModDesc = 412 Or tnk.TankModDesc = 415 Or tnk.TankModDesc = 475) Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "3.5.3"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If (tnk.TankStatus = 424 Or tnk.TankStatus = 429 Or tnk.TankStatus = 425) And (tnk.TankModDesc = 476) And (Date.Compare(tnk.LinedInteriorInstallDate, DateAdd(DateInterval.Year, -10, Now.Date)) < 0) Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "3.5.4"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If (tnk.TankStatus = 424 Or tnk.TankStatus = 429 Or tnk.TankStatus = 425) And (tnk.TankModDesc = 412 Or tnk.TankModDesc = 415 Or tnk.TankModDesc = 475) Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "3.6.1"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If (tnk.TankStatus = 424 Or tnk.TankStatus = 429 Or tnk.TankStatus = 425) And (tnk.pipesCollection.Count > 0) Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "3.6.2"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If (tnk.TankStatus = 424 Or tnk.TankStatus = 429 Or tnk.TankStatus = 425) And (pipe.PipeModDesc = 260 Or pipe.PipeModDesc = 263) Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "3.6.3"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If (tnk.TankStatus = 424 Or tnk.TankStatus = 429 Or tnk.TankStatus = 425) And (pipe.PipeModDesc = 260 Or pipe.PipeModDesc = 263) Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "3.7.1"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeTermArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeTermArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipeTerm And Not (pipe Is Nothing) Then
                                        If (tnk.TankStatus = 424 Or tnk.TankStatus = 429 Or tnk.TankStatus = 425) And tnk.pipesCollection.Count > 0 And Not (pipe.PipeMatDesc = 250 Or pipe.PipeMatDesc = 251) Then
                                            If Not oInspCLInfo.PipeTermArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeTermArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "3.7.2"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeTermArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeTermArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipeTerm And Not (pipe Is Nothing) Then
                                        If (tnk.TankStatus = 424 Or tnk.TankStatus = 429 Or tnk.TankStatus = 425) And tnk.pipesCollection.Count > 0 And Not (pipe.PipeMatDesc = 250 Or pipe.PipeMatDesc = 251) Then
                                            If Not oInspCLInfo.PipeTermArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeTermArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "3.7.3"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeTermArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeTermArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipeTerm And Not (pipe Is Nothing) Then
                                        If (tnk.TankStatus = 424 Or tnk.TankStatus = 429 Or tnk.TankStatus = 425) And (pipe.TermTypeDisp = 490 Or pipe.TermTypeTank = 483) Then
                                            If Not oInspCLInfo.PipeTermArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeTermArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "3.7.4"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeTermArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeTermArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipeTerm And Not (pipe Is Nothing) Then
                                        If (tnk.TankStatus = 424 Or tnk.TankStatus = 429 Or tnk.TankStatus = 425) Then' And (pipe.TermTypeDisp = 489 Or pipe.TermTypeTank = 482) Then
                                            If Not oInspCLInfo.PipeTermArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeTermArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "3.7.5"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeTermArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeTermArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipeTerm And Not (pipe Is Nothing) Then
                                        If (tnk.TankStatus = 424 Or tnk.TankStatus = 429 Or tnk.TankStatus = 425) And (pipe.TermTypeDisp = 611 Or pipe.TermTypeTank = 610) Then
                                            If Not oInspCLInfo.PipeTermArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeTermArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "3.7.6"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeTermArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeTermArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipeTerm And Not (pipe Is Nothing) Then
                                        If (tnk.TankStatus = 424 Or tnk.TankStatus = 429 Or tnk.TankStatus = 425) And (pipe.TermTypeDisp = 611 Or pipe.TermTypeTank = 610) Then
                                            If Not oInspCLInfo.PipeTermArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeTermArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "4.1"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And Not tnk.TankEmergen Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "4.2.1"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.TankLD = 335 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "4.2.2"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.TankLD = 335 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "4.2.3"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.TankLD = 335 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "4.2.4"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.TankLD = 335 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "4.2.5"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.TankLD = 335 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "4.2.6"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.TankLD = 335 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "4.2.7"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.TankLD = 335 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "4.2.8"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.TankLD = 335 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "4.3.1"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.TankLD = 338 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "4.3.2"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.TankLD = 338 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "4.3.3"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.TankLD = 338 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "4.3.4"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.TankLD = 338 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "4.3.5"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.TankLD = 338 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "4.3.6"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.TankLD = 338 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "4.4.1"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.TankLD = 336 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "4.4.2"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.TankLD = 336 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "4.4.3"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.TankLD = 336 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "4.4.4", "4.4.5", "4.4.5.1"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.TankLD = 336 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "4.4.6"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.TankLD = 336 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "4.5.1"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.TankLD = 340 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "4.5.2"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.TankLD = 340 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "4.6.1"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.TankLD = 337 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "4.6.2"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.TankLD = 337 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "4.6.3"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.TankLD = 337 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "4.6.4"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.TankLD = 337 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "4.7.1"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.TankLD = 343 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "4.7.2"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.TankLD = 343 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "4.8.1"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.TankLD = 339 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "4.8.2"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.TankLD = 339 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "4.8.3", "4.8.4", "4.8.4.1"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.TankLD = 339 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                    ' MANJU 3/31/06 START
                                Case "4.8.5"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 424 And tnk.TankLD = 339 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "5.1"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And Not tnk.TankEmergen And Not pipe.PipeTypeDesc = 267 And Not pipe.PipeTypeDesc = 0 Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "5.2.1"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And pipe.PipeLD = 241 Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "5.2.2"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And pipe.PipeLD = 241 Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "5.2.3"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And pipe.PipeLD = 241 Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "5.2.4"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And pipe.PipeLD = 241 Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                    ' MANJU 3/31/06 END
                                Case "5.2.8"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And pipe.PipeLD = 241 Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "5.3.1"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And pipe.PipeLD = 245 And pipe.PipeTypeDesc = 266 Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "5.3.2"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And pipe.PipeLD = 245 And pipe.PipeTypeDesc = 268 Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "5.3.3"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And pipe.PipeLD = 245 Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "5.4.1"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And pipe.PipeLD = 246 Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "5.4.2"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And pipe.PipeLD = 246 Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "5.5.1"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And pipe.PipeLD = 244 Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "5.5.2"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And pipe.PipeLD = 244 Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "5.6.1"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And (pipe.PipeLD = 242 Or pipe.PipeLD = 243) Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "5.6.2"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And (pipe.PipeLD = 242 Or pipe.PipeLD = 243) Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "5.6.3"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And (pipe.PipeLD = 242 Or pipe.PipeLD = 243) Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "5.6.4"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And (pipe.PipeLD = 242 Or pipe.PipeLD = 243) Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "5.7.1", "5.7.2", "5.7.2.1"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And pipe.PipeLD = 242 Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "5.8.1"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And pipe.PipeLD = 243 Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "5.8.2"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And pipe.PipeLD = 243 Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "5.8.3", "5.8.4", "5.8.4.1"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And pipe.PipeLD = 243 Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "5.8.5"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And pipe.PipeLD = 243 Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "5.9.1"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And pipe.PipeTypeDesc = 266 And (Not tnk.TankEmergen Or pipe.PipeLD <> 243) Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "5.9.2"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And pipe.PipeTypeDesc = 266 And (Not tnk.TankEmergen Or pipe.PipeLD <> 243) Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "5.9.3"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And pipe.PipeTypeDesc = 266 And (Not tnk.TankEmergen Or pipe.PipeLD <> 243) Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "5.9.4"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And pipe.PipeTypeDesc = 266 And (Not tnk.TankEmergen Or pipe.PipeLD <> 243) Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "5.9.5"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And pipe.PipeTypeDesc = 266 And (Not tnk.TankEmergen Or pipe.PipeLD <> 243) Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "5.9.6"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And pipe.PipeLD = 246 Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "6.1"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 Then
                                            Dim bolAddPipeID As Boolean = False
                                            For Each comp As MUSTER.Info.CompartmentInfo In tnk.CompartmentCollection.Values
                                                If pipe.CompartmentNumber = comp.COMPARTMENTNumber And comp.Substance <> 314 Then
                                                    bolAddPipeID = True
                                                    Exit For
                                                End If
                                                'If comp.Substance <> 314 Then
                                                '    bolAddPipeID = True
                                                '    Exit For
                                                'End If
                                            Next
                                            If Not tnk.TankEmergen And bolAddPipeID Then
                                                If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                    oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                                End If
                                            End If
                                        End If
                                    End If
                                Case "6.2.1"
                                    If oInspCLInfo.AppliesToPipeTerm And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 AndAlso pipe.TermTypeDisp = 489 AndAlso Date.Compare(pipe.PipeInstallDate, CDate("9/30/08")) > 0 Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "6.2"
                                    If oInspCLInfo.AppliesToPipeTerm AndAlso Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 AndAlso pipe.TermTypeDisp = 489 Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If

                                Case "6.3"
                                    If oInspCLInfo.AppliesToPipe AndAlso Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 AndAlso pipe.PipeTypeDesc = 266 Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "6.4.1"
                                    If oInspCLInfo.AppliesToPipeTerm AndAlso Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 AndAlso pipe.TermTypeTank = 482 AndAlso Date.Compare(pipe.PipeInstallDate, CDate("9/30/08")) > 0 Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "6.4"
                                    If oInspCLInfo.AppliesToPipeTerm And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 AndAlso pipe.TermTypeTank = 482 Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "6.5"
                                    'If Not (pipe Is Nothing) Then
                                    '    If oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                    '        oInspCLInfo.PipeArrayList.Remove(pipe.PipeID)
                                    '        'If colPipeIDIndex.Contains(pipe.PipeID.ToString) Then
                                    '        '    colPipeIDIndex.Remove(pipe.PipeID.ToString)
                                    '        'End If
                                    '    End If
                                    'End If
                                    If oInspCLInfo.AppliesToPipe And Not (pipe Is Nothing) Then
                                        If tnk.TankStatus = 424 And (pipe.PipeMatDesc = 252 Or pipe.PipeMatDesc = 254 Or pipe.PipeMatDesc = 494) Then
                                            If Not oInspCLInfo.PipeArrayList.Contains(pipe.PipeID) Then
                                                oInspCLInfo.PipeArrayList.Add(pipe.PipeID)
                                            End If
                                        End If
                                    End If
                                Case "7.1"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 425 Or tnk.TankStatus = 429 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "7.2"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 425 Or tnk.TankStatus = 429 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                                Case "7.3"
                                    'If oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                    '    oInspCLInfo.TankArrayList.Remove(tnk.TankId)
                                    'End If
                                    If oInspCLInfo.AppliesToTank Then
                                        If tnk.TankStatus = 425 Or tnk.TankStatus = 429 Then
                                            If Not oInspCLInfo.TankArrayList.Contains(tnk.TankId) Then
                                                oInspCLInfo.TankArrayList.Add(tnk.TankId)
                                            End If
                                        End If
                                    End If
                            End Select
                        End If
                    Next
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function ValidateData() As Boolean
            Try
                'Dim facSave As Boolean = oOwner.Facilities.Save()
                Return True
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetCLInspectionHistory(Optional ByVal [readOnly] As Boolean = False) As DataSet
            Dim ds As DataSet
            Try
                ds = oInspectionChecklistMasterDB.DBGetCLInspectionHistory(oInspection.ID)
                If [readOnly] Then
                    For Each dt As DataTable In ds.Tables
                        For Each col As DataColumn In dt.Columns
                            col.ReadOnly = True
                        Next
                    Next
                    'Else
                    '    If ds.Tables(0).Rows.Count = 0 Then
                    '        Dim dr As DataRow
                    '        dr = ds.Tables(0).NewRow
                    '        dr("INS_DATES_ID") = 0
                    '        dr("INSPECTION_ID") = oInspection.ID
                    '        dr("DEQ INSPECTOR") = DBNull.Value
                    '        If Date.Compare(oInspection.RescheduledDate, CDate("01/01/0001")) = 0 Then
                    '            dr("DATE INSPECTED") = oInspection.ScheduledDate
                    '            dr("TIME IN") = oInspection.ScheduledTime
                    '        Else
                    '            dr("DATE INSPECTED") = oInspection.RescheduledDate
                    '            dr("TIME IN") = oInspection.RescheduledTime
                    '        End If
                    '        dr("TIME OUT") = DBNull.Value
                    '        dr("DELETED") = False
                    '        ds.Tables(0).Rows.Add(dr)
                    'End If
                End If
                Return ds
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Sub PutCLInspectionHistory(ByRef id As Int64, ByVal inspectionID As Int64, ByVal staffID As Int64, ByVal insp_Date As Date, ByVal timeIn As String, ByVal timeOut As String, ByVal bolDeleted As Boolean, ByVal moduleID As Integer, ByVal UserID As String, ByRef returnVal As String, ByVal staffIDForSecurity As Integer)
            Try
                oInspectionChecklistMasterDB.DBPutCLInspectionHistory(id, inspectionID, staffID, insp_Date, timeIn, timeOut, bolDeleted, moduleID, UserID, returnVal, staffIDForSecurity)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Private Sub MarkResponsesDeleted(ByVal bolDeleted As Boolean)
            Try
                For Each response As MUSTER.Info.InspectionResponsesInfo In oInspection.ResponsesCollection.Values
                    response.Deleted = bolDeleted
                Next
                'For Each cpreading As MUSTER.Info.InspectionCPReadingsInfo In oInspection.CPReadingsCollection.Values
                '    cpreading.Deleted = bolDeleted
                'Next
                For Each mwell As MUSTER.Info.InspectionMonitorWellsInfo In oInspection.MonitorWellsCollection.Values
                    mwell.Deleted = bolDeleted
                Next
                For Each ccat As MUSTER.Info.InspectionCCATInfo In oInspection.CCATsCollection.Values
                    ccat.Deleted = bolDeleted
                Next
                For Each citation As MUSTER.Info.InspectionCitationInfo In oInspection.CitationsCollection.Values
                    citation.Deleted = bolDeleted
                Next
                For Each discrep As MUSTER.Info.InspectionDiscrepInfo In oInspection.DiscrepsCollection.Values
                    discrep.Deleted = bolDeleted
                Next
                For Each rect As MUSTER.Info.InspectionRectifierInfo In oInspection.RectifiersCollection.Values
                    rect.Deleted = bolDeleted
                Next
                'For Each sketch As MUSTER.Info.InspectionSketchInfo In oInspection.SketchsCollection.Values
                '    sketch.Deleted = bolDeleted
                'Next
                For Each soc As MUSTER.Info.InspectionSOCInfo In oInspection.SOCsCollection.Values
                    soc.Deleted = True
                Next
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "Collection Operations"
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal id As Int64)
            Try
                oInspectionChecklistMasterInfo = oInspectionChecklistMasterDB.DBGetByID(id)
                If oInspectionChecklistMasterInfo.ID = 0 Then
                    oInspectionChecklistMasterInfo.ID = nID
                    nID -= 1
                End If
                oInspection.ChecklistMasterCollection.Add(oInspectionChecklistMasterInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oInspectionChecklistMaster As MUSTER.Info.InspectionChecklistMasterInfo)
            Try
                oInspectionChecklistMasterInfo = oInspectionChecklistMaster
                If oInspectionChecklistMasterInfo.ID = 0 Then
                    oInspectionChecklistMasterInfo.ID = nID
                    nID -= 1
                End If
                oInspection.ChecklistMasterCollection.Add(oInspectionChecklistMasterInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal id As Int64)
            Try
                If oInspection.ChecklistMasterCollection.Contains(id) Then
                    oInspection.ChecklistMasterCollection.Remove(id)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oInspectionChecklistMaster As MUSTER.Info.InspectionChecklistMasterInfo)
            Try
                If oInspection.ChecklistMasterCollection.Contains(oInspectionChecklistMaster) Then
                    oInspection.ChecklistMasterCollection.Remove(oInspectionChecklistMaster)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByVal UserID As String, ByRef returnVal As String, Optional ByVal strUser As String = "", Optional ByVal NormalfacilityFlush As Boolean = true)
            Try


                If oOwner.Facilities.ID > 0 Then
                    oOwner.Facilities.Flush(moduleID, staffID, returnVal, strUser, Not NormalfacilityFlush)
                End If
                oInspectionResponses.Flush(moduleID, staffID, returnVal)
                If GetCLInspectionHistory(True).Tables(0).Rows.Count = 0 Then
                    If Date.Compare(oInspection.RescheduledDate, CDate("01/01/0001")) = 0 Then
                        PutCLInspectionHistory(0, oInspection.ID, oInspection.StaffID, oInspection.ScheduledDate, oInspection.ScheduledTime, String.Empty, False, moduleID, UserID, returnVal, staffID)
                    Else
                        PutCLInspectionHistory(0, oInspection.ID, oInspection.StaffID, oInspection.RescheduledDate, oInspection.RescheduledTime, String.Empty, False, moduleID, UserID, returnVal, staffID)
                    End If
                End If
                oInspectionCCAT.Flush(moduleID, staffID, returnVal)
                oInspectionCitation.Flush(moduleID, staffID, returnVal)
                oInspectionDiscrep.Flush(moduleID, staffID, returnVal)
                oInspectionCPReadings.Flush(moduleID, staffID, returnVal)
                oInspectionMonitorWells.Flush(moduleID, staffID, returnVal)
                oInspectionRectifier.Flush(moduleID, staffID, returnVal)
                oInspectionSketch.Flush(moduleID, staffID, returnVal)
                oInspectionComments.Flush(moduleID, staffID, returnVal)
                oInspectionSOC.Flush(moduleID, staffID, returnVal)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "General Operations"
        Public Sub Clear()
            oInspectionChecklistMasterInfo = New MUSTER.Info.InspectionChecklistMasterInfo
        End Sub
        Public Sub Reset()
            oInspectionChecklistMasterInfo.Reset()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function EntityTable() As DataTable
            Dim oInspectionChecklistMasterInfoLocal As New MUSTER.Info.InspectionChecklistMasterInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Try
                tbEntityTable.Columns.Add("ID")
                tbEntityTable.Columns.Add("CheckListItemNumber")
                tbEntityTable.Columns.Add("SOC")
                tbEntityTable.Columns.Add("Header")
                tbEntityTable.Columns.Add("Header Question Text")
                'tbEntityTable.Columns.Add("Response Table")
                tbEntityTable.Columns.Add("Applies To")
                tbEntityTable.Columns.Add("Citation")
                tbEntityTable.Columns.Add("When Visible")
                tbEntityTable.Columns.Add("CCAT")
                tbEntityTable.Columns.Add("Deleted")
                tbEntityTable.Columns.Add("Created By")
                tbEntityTable.Columns.Add("Date Created")
                tbEntityTable.Columns.Add("Last Edited By")
                tbEntityTable.Columns.Add("Date Last Edited")

                For Each oInspectionChecklistMasterInfoLocal In oInspection.ChecklistMasterCollection.Values
                    dr = tbEntityTable.NewRow()
                    dr("ID") = oInspectionChecklistMasterInfoLocal.ID
                    dr("CheckListItemNumber") = oInspectionChecklistMasterInfoLocal.CheckListItemNumber
                    dr("SOC") = oInspectionChecklistMasterInfoLocal.SOC
                    dr("Header") = oInspectionChecklistMasterInfoLocal.Header
                    dr("Header Question Text") = oInspectionChecklistMasterInfoLocal.HeaderQuestionText
                    'dr("Response Table") = oInspectionChecklistMasterInfoLocal.ResponseTable
                    dr("Applies To Tank") = oInspectionChecklistMasterInfoLocal.AppliesToTank
                    dr("Applies To Pipe") = oInspectionChecklistMasterInfoLocal.AppliesToPipe
                    dr("Applies To PipeTerm") = oInspectionChecklistMasterInfoLocal.AppliesToPipeTerm
                    dr("Citation") = oInspectionChecklistMasterInfoLocal.Citation
                    dr("When Visible") = oInspectionChecklistMasterInfoLocal.WhenVisible
                    dr("CCAT") = oInspectionChecklistMasterInfoLocal.CCAT
                    dr("Deleted") = oInspectionChecklistMasterInfoLocal.Deleted
                    dr("Created By") = oInspectionChecklistMasterInfoLocal.CreatedBy
                    dr("Date Created") = oInspectionChecklistMasterInfoLocal.CreatedOn
                    dr("Last Edited By") = oInspectionChecklistMasterInfoLocal.ModifiedBy
                    dr("Date Last Edited") = oInspectionChecklistMasterInfoLocal.ModifiedOn
                    tbEntityTable.Rows.Add(dr)
                Next
                Return tbEntityTable
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function TanksPipesTables(ByVal facID As Int64, ByVal inspectionID As Integer) As DataSet
            Try
                Dim strSQL As String
                If oInspection.SubmittedDate.ToString Is DBNull.Value Or oInspection.SubmittedDate = Date.MinValue Or oInspection.SubmittedDate.ToString = String.Empty Then
                    strSQL = "SELECT * FROM vINSPECTION_TANK_DISPLAY_DATA WHERE FACILITY_ID = " + facID.ToString + " ORDER BY POSITION, [TANK #], COMPARTMENT_NUMBER;" + _
                    "SELECT * FROM vINSPECTION_PIPES_DISPLAY_DATA WHERE FACILITY_ID = " + facID.ToString + "  AND PARENT_PIPE_ID = 0 ORDER BY POSITION, [PIPE #];" + _
                    "SELECT * FROM vINSPECTION_TERMINATIONS_DISPLAY_DATA WHERE FACILITY_ID = " + facID.ToString + "AND PARENT_PIPE_ID = 0 ORDER BY POSITION, [PIPE #];" + _
                    "SELECT * FROM vINSPECTION_PIPES_DISPLAY_DATA WHERE FACILITY_ID = " + facID.ToString + "  AND PARENT_PIPE_ID > 0 ORDER BY POSITION, [PIPE #];" + _
                    "SELECT * FROM vINSPECTION_TERMINATIONS_DISPLAY_DATA WHERE FACILITY_ID = " + facID.ToString + "AND PARENT_PIPE_ID > 0 ORDER BY POSITION, [PIPE #];"

                    strSQL = String.Format("exec spUpdateTanksPipesToInspectionView {0}; {1}", facID, strSQL)

                Else
                    strSQL = "SELECT * FROM vINSPECTION_TANK_DISPLAY_ARCHIVE_DATA WHERE FACILITY_ID = " + facID.ToString + " ORDER BY POSITION, [TANK #], COMPARTMENT_NUMBER;" + _
                    "SELECT * FROM vINSPECTION_PIPES_DISPLAY_ARCHIVE_DATA WHERE FACILITY_ID = " + facID.ToString + " AND PARENT_PIPE_ID = 0 ORDER BY POSITION, [PIPE #];" + _
                    "SELECT * FROM vINSPECTION_TERMINATIONS_DISPLAY_ARCHIVE_DATA WHERE FACILITY_ID = " + facID.ToString + " AND PARENT_PIPE_ID = 0 ORDER BY POSITION, [PIPE #];" + _
                    "SELECT * FROM vINSPECTION_PIPES_DISPLAY_ARCHIVE_DATA WHERE FACILITY_ID = " + facID.ToString + "  AND PARENT_PIPE_ID > 0 ORDER BY POSITION, [PIPE #];" + _
                    "SELECT * FROM vINSPECTION_TERMINATIONS_DISPLAY_ARCHIVE_DATA WHERE FACILITY_ID = " + facID.ToString + "AND PARENT_PIPE_ID > 0 ORDER BY POSITION, [PIPE #];"
                    strSQL = strSQL.Replace(" WHERE ", String.Format(" WHERE INSPECTIOn_ID = {0} AND ", inspectionID))

                End If


                Dim ds As DataSet = oInspectionChecklistMasterDB.DBGetDS(strSQL)
                Dim Rel, Rel2 As DataRelation
                Dim c1() As DataColumn = {ds.Tables(1).Columns("PIPE_ID"), ds.Tables(1).Columns("TANK_ID")}
                Dim c2() As DataColumn = {ds.Tables(3).Columns("PARENT_PIPE_ID"), ds.Tables(3).Columns("TANK_ID")}

                Dim c3() As DataColumn = {ds.Tables(2).Columns("PIPE_ID"), ds.Tables(2).Columns("TANK_ID")}
                Dim c4() As DataColumn = {ds.Tables(4).Columns("PARENT_PIPE_ID"), ds.Tables(4).Columns("TANK_ID")}

                Rel = New DataRelation("PipeToExt", c1, c2, False)
                Rel2 = New DataRelation("TermPipeToExt", c3, c4, False)

                ds.Relations.Add(Rel)
                ds.Relations.Add(Rel2)

                Return ds
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Private Function GetCCAT(ByVal checkList As MUSTER.Info.InspectionChecklistMasterInfo) As String
            Try
                If checkList.Citation <> -1 Then
                    For Each citation As MUSTER.Info.InspectionCitationInfo In oInspection.CitationsCollection.Values
                        If citation.QuestionID = checkList.ID Then
                            Return citation.CCAT
                        End If
                    Next
                End If
                Return String.Empty
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function


        Public Function AddUserDefineRows(ByVal gridNum As Integer, ByVal id As Integer) As DataTable
            Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo
            Dim citation As MUSTER.Info.InspectionCitationInfo

            Try

                Dim dr As DataSet = Me.oInspectionChecklistMasterDB.DBGetDS(String.Format("exec spGetInspectionUserDefinedCheckList null,{0},null,null,{1}", _
                                                                              gridNum, id))
                Return dr.Tables(0)

            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try


        End Function

        Public Function RegTable(Optional ByVal [readOnly] As Boolean = False) As DataTable
            Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo
            Dim citation As MUSTER.Info.InspectionCitationInfo
            Dim dr As DataRow
            Dim tb As New DataTable
            Try
                tb.Columns.Add("CL_POSITION", GetType(Int64))
                tb.Columns.Add("Line#")
                tb.Columns.Add("Question")
                tb.Columns.Add("Yes", GetType(Boolean))
                tb.Columns.Add("No", GetType(Boolean))
                tb.Columns.Add("N/A", GetType(Boolean))
                tb.Columns.Add("CCAT")
                tb.Columns.Add("ID")
                tb.Columns.Add("INSPECTION_ID")
                tb.Columns.Add("QUESTION_ID")
                tb.Columns.Add("SOC")
                tb.Columns.Add("RESPONSE")
                tb.Columns.Add("HEADER")
                tb.Columns.Add("CITATION")
                tb.Columns.Add("FORE_COLOR")
                tb.Columns.Add("BACK_COLOR")

                For Each response As MUSTER.Info.InspectionResponsesInfo In oInspection.ResponsesCollection.Values
                    If Not response.Deleted Then
                        checkList = oInspection.ChecklistMasterCollection.Item(response.QuestionID)
                        If Not checkList Is Nothing Then
                            If (checkList.CheckListItemNumber = "1" Or checkList.CheckListItemNumber.StartsWith("1.")) And checkList.Show Then
                                dr = tb.NewRow()
                                dr("CL_POSITION") = checkList.Position
                                dr("Line#") = checkList.CheckListItemNumber
                                dr("Question") = checkList.HeaderQuestionText
                                If checkList.Header Then
                                    dr("Yes") = False
                                    dr("No") = False
                                    dr("N/A") = False
                                Else
                                    dr("Yes") = IIf(response.Response = 1 Or response.Response = -2, True, False)
                                    dr("No") = IIf(response.Response = 0 Or response.Response = -2, True, False)
                                    dr("N/A") = IIf(response.Response = 2 Or response.Response = -2, True, False)
                                End If
                                dr("CCAT") = GetCCAT(checkList)
                                dr("ID") = response.ID
                                dr("INSPECTION_ID") = response.InspectionID
                                dr("QUESTION_ID") = response.QuestionID
                                dr("SOC") = response.SOC
                                dr("RESPONSE") = response.Response
                                dr("HEADER") = checkList.Header
                                dr("CITATION") = checkList.CCAT
                                dr("FORE_COLOR") = checkList.ForeColor
                                dr("BACK_COLOR") = checkList.BackColor
                                tb.Rows.Add(dr)
                            End If
                        End If
                    End If
                Next
                tb.DefaultView.Sort = "CL_POSITION"
                If [readOnly] Then
                    For Each col As DataColumn In tb.Columns
                        col.ReadOnly = True
                    Next
                End If


                Return tb

            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function SpillTable(Optional ByVal [readOnly] As Boolean = False) As DataTable
            Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo
            Dim dr As DataRow
            Dim tb As New DataTable
            Try
                tb.Columns.Add("CL_POSITION", GetType(Int64))
                tb.Columns.Add("Line#")
                tb.Columns.Add("Question")
                tb.Columns.Add("Yes", GetType(Boolean))
                tb.Columns.Add("No", GetType(Boolean))
                tb.Columns.Add("N/A", GetType(Boolean))

                tb.Columns.Add("CCAT")
                tb.Columns.Add("ID")
                tb.Columns.Add("INSPECTION_ID")
                tb.Columns.Add("QUESTION_ID")
                tb.Columns.Add("SOC")
                tb.Columns.Add("RESPONSE")
                tb.Columns.Add("HEADER")
                tb.Columns.Add("CITATION")
                tb.Columns.Add("FORE_COLOR")
                tb.Columns.Add("BACK_COLOR")

                For Each response As MUSTER.Info.InspectionResponsesInfo In oInspection.ResponsesCollection.Values
                    If Not response.Deleted Then
                        checkList = oInspection.ChecklistMasterCollection.Item(response.QuestionID)
                        If Not checkList Is Nothing Then
                            If checkList.CheckListItemNumber >= "2" And checkList.CheckListItemNumber < "3" And checkList.Show Then
                                dr = tb.NewRow()
                                dr("CL_POSITION") = checkList.Position
                                dr("Line#") = checkList.CheckListItemNumber
                                dr("Question") = checkList.HeaderQuestionText
                                dr("Yes") = IIf(response.Response = 1 Or response.Response = -2, True, False)
                                dr("No") = IIf(response.Response = 0 Or response.Response = -2, True, False)
                                dr("N/A") = IIf(response.Response = 2 Or response.Response = -2, True, False)
                                dr("CCAT") = GetCCAT(checkList)
                                dr("ID") = response.ID
                                dr("INSPECTION_ID") = response.InspectionID
                                dr("QUESTION_ID") = response.QuestionID
                                dr("SOC") = response.SOC
                                dr("RESPONSE") = response.Response
                                dr("HEADER") = checkList.Header
                                dr("CITATION") = checkList.CCAT
                                dr("FORE_COLOR") = checkList.ForeColor
                                dr("BACK_COLOR") = checkList.BackColor
                                tb.Rows.Add(dr)
                            End If
                        End If
                    End If
                Next
                tb.DefaultView.Sort = "CL_POSITION"
                If [readOnly] Then
                    For Each col As DataColumn In tb.Columns
                        col.ReadOnly = True
                    Next
                End If
                Return tb
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function


        Public Function OtherQusetionsTable(Optional ByVal [readOnly] As Boolean = False) As DataTable
            Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo
            Dim dr As DataRow
            Dim tb As New DataTable
            Try
                tb.Columns.Add("CL_POSITION", GetType(Int64))
                tb.Columns.Add("Line#")
                tb.Columns.Add("Question")
                tb.Columns.Add("Yes", GetType(Boolean))
                tb.Columns.Add("No", GetType(Boolean))
                tb.Columns.Add("N/A", GetType(Boolean))

                tb.Columns.Add("CCAT")
                tb.Columns.Add("ID")
                tb.Columns.Add("INSPECTION_ID")
                tb.Columns.Add("QUESTION_ID")
                tb.Columns.Add("SOC")
                tb.Columns.Add("RESPONSE")
                tb.Columns.Add("HEADER")
                tb.Columns.Add("CITATION")
                tb.Columns.Add("FORE_COLOR")
                tb.Columns.Add("BACK_COLOR")

                Dim answered As Boolean = True 'False 'False - turn 12.OPER off; True - turn 12.OPER on

                For Each response As MUSTER.Info.InspectionResponsesInfo In oInspection.ResponsesCollection.Values
                    If Not response.Deleted Then
                        checkList = oInspection.ChecklistMasterCollection.Item(response.QuestionID)
                        If Not checkList Is Nothing Then

                            If checkList.CheckListItemNumber = "1.13" AndAlso response.Response <> 0 Then
                                answered = True
                            End If

                            If checkList.CheckListItemNumber >= "12" And checkList.CheckListItemNumber < "13" And answered And Me.hasOwnerAsDesignatedOperator Then
                                dr = tb.NewRow()
                                dr("CL_POSITION") = checkList.Position
                                dr("Line#") = checkList.CheckListItemNumber
                                dr("Question") = checkList.HeaderQuestionText

                                If checkList.Header Then
                                    dr("Yes") = False
                                    dr("No") = False
                                    dr("N/A") = False
                                Else
                                    dr("Yes") = IIf(response.Response = 1 Or response.Response = -2, True, False)
                                    dr("No") = IIf(response.Response = 0 Or response.Response = -2, True, False)
                                    dr("N/A") = IIf(response.Response = 2 Or response.Response = -2, True, False)
                                End If
                                'dr("Yes") = IIf(response.Response = 1 Or response.Response = -2, True, False)
                                ' dr("No") = IIf(response.Response = 0 Or response.Response = -2, True, False)
                                ' dr("N/A") = IIf(response.Response = 2 Or response.Response = -2, True, False)
                                dr("CCAT") = GetCCAT(checkList)
                                dr("ID") = response.ID
                                dr("INSPECTION_ID") = response.InspectionID
                                dr("QUESTION_ID") = response.QuestionID
                                dr("SOC") = response.SOC
                                dr("RESPONSE") = response.Response
                                dr("HEADER") = checkList.Header
                                dr("CITATION") = checkList.CCAT
                                dr("FORE_COLOR") = checkList.ForeColor
                                dr("BACK_COLOR") = checkList.BackColor
                                tb.Rows.Add(dr)
                            End If
                        End If
                    End If
                Next
                tb.DefaultView.Sort = "CL_POSITION"
                If [readOnly] Then
                    For Each col As DataColumn In tb.Columns
                        col.ReadOnly = True
                    Next
                End If
                Return tb
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function CPTable(Optional ByVal [readOnly] As Boolean = False) As DataSet
            Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo
            Dim dsRel As DataRelation
            Dim ds As New DataSet
            Dim dr As DataRow
            Dim tb As New DataTable
            Dim tbCPRect As New DataTable
            Dim tbCPTank As New DataTable
            Dim tbCPTankRemote As New DataTable
            Dim tbCPTankGalIC As New DataTable
            Dim tbCPTankInspectorTested As New DataTable
            Dim tbCPPipe As New DataTable
            Dim tbCPPipeRemote As New DataTable
            Dim tbCPPipeGalIC As New DataTable
            Dim tbCPPipeInspectorTested As New DataTable
            Dim tbCPTerm As New DataTable
            Dim tbCPTermRemote As New DataTable
            Dim tbCPTermGalIC As New DataTable
            Dim tbCPTermInspectorTested As New DataTable
            Try
                tb.Columns.Add("CL_POSITION", GetType(Int64))
                tb.Columns.Add("Line#")
                tb.Columns.Add("Question")
                tb.Columns.Add("Yes", GetType(Boolean))
                tb.Columns.Add("No", GetType(Boolean))
                tb.Columns.Add("N/A", GetType(Boolean))
                tb.Columns.Add("CCAT")
                tb.Columns.Add("ID")
                tb.Columns.Add("INSPECTION_ID")
                tb.Columns.Add("QUESTION_ID")
                tb.Columns.Add("SOC")
                tb.Columns.Add("RESPONSE")
                tb.Columns.Add("HEADER")
                tb.Columns.Add("CITATION")
                tb.Columns.Add("FORE_COLOR")
                tb.Columns.Add("BACK_COLOR")

                For Each response As MUSTER.Info.InspectionResponsesInfo In oInspection.ResponsesCollection.Values
                    If Not response.Deleted Then
                        checkList = oInspection.ChecklistMasterCollection.Item(response.QuestionID)
                        If Not checkList Is Nothing Then
                            If checkList.CheckListItemNumber >= "3" And checkList.CheckListItemNumber < "4" And checkList.Show Then
                                dr = tb.NewRow()
                                dr("CL_POSITION") = checkList.Position
                                dr("Line#") = checkList.CheckListItemNumber
                                dr("Question") = checkList.HeaderQuestionText
                                dr("Yes") = IIf(response.Response = 1 Or response.Response = -2, True, False)
                                dr("No") = IIf(response.Response = 0 Or response.Response = -2, True, False)
                                dr("N/A") = IIf(response.Response = 2 Or response.Response = -2, True, False)
                                dr("CCAT") = GetCCAT(checkList)
                                dr("ID") = response.ID
                                dr("INSPECTION_ID") = response.InspectionID
                                dr("QUESTION_ID") = response.QuestionID
                                dr("SOC") = response.SOC
                                dr("RESPONSE") = response.Response
                                dr("HEADER") = checkList.Header
                                dr("CITATION") = checkList.CCAT
                                dr("FORE_COLOR") = checkList.ForeColor
                                dr("BACK_COLOR") = checkList.BackColor
                                tb.Rows.Add(dr)
                            End If
                        End If
                    End If
                Next

                tbCPRect.Columns.Add("Volts", GetType(Double))
                tbCPRect.Columns.Add("Amps", GetType(Double))
                tbCPRect.Columns.Add("Hours", GetType(Double))
                tbCPRect.Columns.Add("How Long", GetType(String))
                tbCPRect.Columns.Add("ID")
                tbCPRect.Columns.Add("INSPECTION_ID")
                tbCPRect.Columns.Add("QUESTION_ID")
                tbCPRect.Columns.Add("RECITIFIER_ON")
                'tbCPRect.Columns.Add("INOP_HOW_LONG", GetType(Long))

                tbCPTank.Columns.Add("Line#")
                tbCPTank.Columns.Add("Tank#", GetType(Int64))
                tbCPTank.Columns.Add("Fuel Type", GetType(String))
                tbCPTank.Columns.Add("Contact Point")
                tbCPTank.Columns.Add("Local Reference Cell Placement")
                tbCPTank.Columns.Add("Local/On")
                tbCPTank.Columns.Add("Remote/Off")
                tbCPTank.Columns.Add("Pass", GetType(Boolean))
                tbCPTank.Columns.Add("Fail", GetType(Boolean))
                tbCPTank.Columns.Add("Incon", GetType(Boolean))
                tbCPTank.Columns.Add("ID")
                tbCPTank.Columns.Add("INSPECTION_ID")
                tbCPTank.Columns.Add("QUESTION_ID")
                tbCPTank.Columns.Add("TANK_PIPE_ID")
                tbCPTank.Columns.Add("TANK_PIPE_ENTITY_ID")
                tbCPTank.Columns.Add("TANK_INDEX", GetType(Integer))
                'tbCPTank.Columns.Add("TANK_DISPENSER")
                'tbCPTank.Columns.Add("GALVANIC")
                'tbCPTank.Columns.Add("IMPRESSED_CURRENT")
                tbCPTank.Columns.Add("PASSFAILINCON")
                tbCPTank.Columns.Add("CITATION")
                tbCPTank.Columns.Add("LINE_NUMBER", GetType(Integer))

                tbCPTankGalIC.Columns.Add("Galvanic", GetType(Boolean))
                tbCPTankGalIC.Columns.Add("Impressed Current", GetType(Boolean))
                tbCPTankGalIC.Columns.Add("GALVANIC_IC_RESPONSE", GetType(Integer))
                tbCPTankGalIC.Columns.Add("ID")
                tbCPTankGalIC.Columns.Add("INSPECTION_ID")
                tbCPTankGalIC.Columns.Add("QUESTION_ID")

                tbCPTankRemote.Columns.Add("Description of Remote Reference Cell Placement", GetType(String))
                tbCPTankRemote.Columns.Add("ID")
                tbCPTankRemote.Columns.Add("INSPECTION_ID")
                tbCPTankRemote.Columns.Add("QUESTION_ID")

                tbCPTankInspectorTested.Columns.Add("ID")
                tbCPTankInspectorTested.Columns.Add("INSPECTION_ID")
                tbCPTankInspectorTested.Columns.Add("QUESTION_ID")
                tbCPTankInspectorTested.Columns.Add("BLANK", GetType(String))
                tbCPTankInspectorTested.Columns.Add("Yes", GetType(Boolean))
                tbCPTankInspectorTested.Columns.Add("No", GetType(Boolean))

                tbCPTankInspectorTested.Columns.Add("TESTED_BY_INSPECTOR_RESPONSE", GetType(Boolean))

                tbCPPipe.Columns.Add("Line#")
                tbCPPipe.Columns.Add("Pipe#", GetType(Int64))
                tbCPPipe.Columns.Add("Fuel Type", GetType(String))
                tbCPPipe.Columns.Add("Contact Point")
                tbCPPipe.Columns.Add("Local Reference Cell Placement")
                tbCPPipe.Columns.Add("Local/On")
                tbCPPipe.Columns.Add("Remote/Off")
                tbCPPipe.Columns.Add("Pass", GetType(Boolean))
                tbCPPipe.Columns.Add("Fail", GetType(Boolean))
                tbCPPipe.Columns.Add("Incon", GetType(Boolean))
                tbCPPipe.Columns.Add("ID")
                tbCPPipe.Columns.Add("INSPECTION_ID")
                tbCPPipe.Columns.Add("QUESTION_ID")
                tbCPPipe.Columns.Add("TANK_PIPE_ID")
                tbCPPipe.Columns.Add("TANK_PIPE_ENTITY_ID")
                tbCPPipe.Columns.Add("PIPE_INDEX", GetType(Integer))
                'tbCPPipe.Columns.Add("TANK_DISPENSER")
                'tbCPPipe.Columns.Add("GALVANIC")
                'tbCPPipe.Columns.Add("IMPRESSED_CURRENT")
                tbCPPipe.Columns.Add("PASSFAILINCON")
                tbCPPipe.Columns.Add("CITATION")
                'tbCPPipe.Columns.Add("Question")
                tbCPPipe.Columns.Add("LINE_NUMBER", GetType(Integer))

                tbCPPipeGalIC.Columns.Add("Galvanic", GetType(Boolean))
                tbCPPipeGalIC.Columns.Add("Impressed Current", GetType(Boolean))
                tbCPPipeGalIC.Columns.Add("GALVANIC_IC_RESPONSE", GetType(Integer))
                tbCPPipeGalIC.Columns.Add("ID")
                tbCPPipeGalIC.Columns.Add("INSPECTION_ID")
                tbCPPipeGalIC.Columns.Add("QUESTION_ID")

                tbCPPipeRemote.Columns.Add("Description of Remote Reference Cell Placement", GetType(String))
                tbCPPipeRemote.Columns.Add("ID")
                tbCPPipeRemote.Columns.Add("INSPECTION_ID")
                tbCPPipeRemote.Columns.Add("QUESTION_ID")

                tbCPPipeInspectorTested.Columns.Add("ID")
                tbCPPipeInspectorTested.Columns.Add("INSPECTION_ID")
                tbCPPipeInspectorTested.Columns.Add("QUESTION_ID")
                tbCPPipeInspectorTested.Columns.Add("BLANK", GetType(String))
                tbCPPipeInspectorTested.Columns.Add("Yes", GetType(Boolean))
                tbCPPipeInspectorTested.Columns.Add("No", GetType(Boolean))
                tbCPPipeInspectorTested.Columns.Add("TESTED_BY_INSPECTOR_RESPONSE", GetType(Boolean))

                tbCPTerm.Columns.Add("Line#")
                tbCPTerm.Columns.Add("Term#", GetType(Int64))
                tbCPTerm.Columns.Add("Fuel Type", GetType(String))
                tbCPTerm.Columns.Add("Contact Point")
                tbCPTerm.Columns.Add("Local Reference Cell Placement")
                tbCPTerm.Columns.Add("Local/On")
                tbCPTerm.Columns.Add("Remote/Off")
                tbCPTerm.Columns.Add("Pass", GetType(Boolean))
                tbCPTerm.Columns.Add("Fail", GetType(Boolean))
                tbCPTerm.Columns.Add("Incon", GetType(Boolean))
                tbCPTerm.Columns.Add("ID")
                tbCPTerm.Columns.Add("INSPECTION_ID")
                tbCPTerm.Columns.Add("QUESTION_ID")
                tbCPTerm.Columns.Add("TANK_PIPE_ID")
                tbCPTerm.Columns.Add("TANK_PIPE_ENTITY_ID")
                tbCPTerm.Columns.Add("TERM_INDEX", GetType(Integer))
                'tbCPTerm.Columns.Add("TANK_DISPENSER")
                'tbCPTerm.Columns.Add("GALVANIC")
                'tbCPTerm.Columns.Add("IMPRESSED_CURRENT")
                tbCPTerm.Columns.Add("PASSFAILINCON")
                tbCPTerm.Columns.Add("CITATION")
                'tbCPTerm.Columns.Add("Question")
                tbCPTerm.Columns.Add("LINE_NUMBER", GetType(Integer))

                tbCPTermGalIC.Columns.Add("Galvanic", GetType(Boolean))
                tbCPTermGalIC.Columns.Add("Impressed Current", GetType(Boolean))
                tbCPTermGalIC.Columns.Add("GALVANIC_IC_RESPONSE", GetType(Integer))
                tbCPTermGalIC.Columns.Add("ID")
                tbCPTermGalIC.Columns.Add("INSPECTION_ID")
                tbCPTermGalIC.Columns.Add("QUESTION_ID")

                tbCPTermRemote.Columns.Add("Description of Remote Reference Cell Placement", GetType(String))
                tbCPTermRemote.Columns.Add("ID")
                tbCPTermRemote.Columns.Add("INSPECTION_ID")
                tbCPTermRemote.Columns.Add("QUESTION_ID")

                tbCPTermInspectorTested.Columns.Add("ID")
                tbCPTermInspectorTested.Columns.Add("INSPECTION_ID")
                tbCPTermInspectorTested.Columns.Add("QUESTION_ID")
                tbCPTermInspectorTested.Columns.Add("BLANK", GetType(String))
                tbCPTermInspectorTested.Columns.Add("Yes", GetType(Boolean))
                tbCPTermInspectorTested.Columns.Add("No", GetType(Boolean))
                tbCPTermInspectorTested.Columns.Add("TESTED_BY_INSPECTOR_RESPONSE", GetType(Boolean))

                For Each rect As MUSTER.Info.InspectionRectifierInfo In oInspection.RectifiersCollection.Values
                    If Not rect.Deleted Then
                        checkList = oInspection.ChecklistMasterCollection.Item(rect.QuestionID)
                        If Not checkList Is Nothing Then
                            If checkList.Show Then
                                dr = tbCPRect.NewRow
                                dr("Volts") = rect.Volts
                                dr("Amps") = rect.Amps
                                dr("Hours") = rect.Hours
                                dr("How Long") = rect.InopHowLong
                                dr("ID") = rect.ID
                                dr("INSPECTION_ID") = rect.InspectionID
                                dr("QUESTION_ID") = rect.QuestionID
                                dr("RECITIFIER_ON") = rect.RectifierOn
                                'dr("INOP_HOW_LONG") = rect.InopHowLong
                                tbCPRect.Rows.Add(dr)
                            End If
                        End If
                    End If
                Next

                For Each cp As MUSTER.Info.InspectionCPReadingsInfo In oInspection.CPReadingsCollection.Values
                    If Not cp.Deleted Then
                        checkList = oInspection.ChecklistMasterCollection.Item(cp.QuestionID)
                        If Not checkList Is Nothing Then
                            If checkList.CheckListItemNumber = "3.5.4" And checkList.Show Then
                                If cp.RemoteReferCellPlacement Then
                                    dr = tbCPTankRemote.NewRow
                                    dr("Description of Remote Reference Cell Placement") = cp.LocalReferCellPlacement
                                    dr("ID") = cp.ID
                                    dr("INSPECTION_ID") = cp.InspectionID
                                    dr("QUESTION_ID") = cp.QuestionID
                                    tbCPTankRemote.Rows.Add(dr)
                                ElseIf cp.GalvanicIC Then
                                    dr = tbCPTankGalIC.NewRow
                                    dr("Galvanic") = IIf(cp.GalvanicICResponse = 0, True, False)
                                    dr("Impressed Current") = IIf(cp.GalvanicICResponse = 1, True, False)
                                    dr("GALVANIC_IC_RESPONSE") = cp.GalvanicICResponse
                                    dr("ID") = cp.ID
                                    dr("INSPECTION_ID") = cp.InspectionID
                                    dr("QUESTION_ID") = cp.QuestionID
                                    tbCPTankGalIC.Rows.Add(dr)
                                ElseIf cp.TestedByInspector Then
                                    dr = tbCPTankInspectorTested.NewRow
                                    dr("BLANK") = String.Empty
                                    dr("Yes") = cp.TestedByInspectorResponse
                                    dr("No") = Not cp.TestedByInspectorResponse
                                    dr("ID") = cp.ID
                                    dr("INSPECTION_ID") = cp.InspectionID
                                    dr("QUESTION_ID") = cp.QuestionID
                                    dr("TESTED_BY_INSPECTOR_RESPONSE") = cp.TestedByInspectorResponse
                                    tbCPTankInspectorTested.Rows.Add(dr)
                                Else
                                    dr = tbCPTank.NewRow
                                    dr("Line#") = "3.5.4." + cp.LineNumber.ToString
                                    dr("Tank#") = cp.TankPipeIndex
                                    dr("Fuel Type") = slTankFuelType.Item(cp.TankPipeID)
                                    dr("Contact Point") = cp.ContactPoint
                                    dr("Local Reference Cell Placement") = cp.LocalReferCellPlacement
                                    dr("Local/On") = cp.LocalOn
                                    dr("Remote/Off") = cp.RemoteOff
                                    dr("Pass") = IIf(cp.PassFailIncon = 1, True, False)
                                    dr("Fail") = IIf(cp.PassFailIncon = 0, True, False)
                                    dr("Incon") = IIf(cp.PassFailIncon = 2, True, False)
                                    dr("ID") = cp.ID
                                    dr("INSPECTION_ID") = cp.InspectionID
                                    dr("QUESTION_ID") = cp.QuestionID
                                    dr("TANK_PIPE_ID") = cp.TankPipeID
                                    dr("TANK_PIPE_ENTITY_ID") = cp.TankPipeEntityID
                                    dr("TANK_INDEX") = cp.TankPipeIndex
                                    'dr("TANK_DISPENSER") = cp.TankDispenser
                                    'dr("GALVANIC") = cp.Galvanic
                                    'dr("IMPRESSED_CURRENT") = cp.ImpressedCurrent
                                    dr("PASSFAILINCON") = cp.PassFailIncon
                                    dr("CITATION") = checkList.CCAT
                                    'dr("Question") = "Description of Remote Reference Cell Placement"
                                    dr("LINE_NUMBER") = cp.LineNumber
                                    tbCPTank.Rows.Add(dr)
                                End If
                            ElseIf checkList.CheckListItemNumber = "3.6.3" And checkList.Show Then
                                If cp.RemoteReferCellPlacement Then
                                    dr = tbCPPipeRemote.NewRow
                                    dr("Description of Remote Reference Cell Placement") = cp.LocalReferCellPlacement
                                    dr("ID") = cp.ID
                                    dr("INSPECTION_ID") = cp.InspectionID
                                    dr("QUESTION_ID") = cp.QuestionID
                                    tbCPPipeRemote.Rows.Add(dr)
                                ElseIf cp.GalvanicIC Then
                                    dr = tbCPPipeGalIC.NewRow
                                    dr("Galvanic") = IIf(cp.GalvanicICResponse = 0, True, False)
                                    dr("Impressed Current") = IIf(cp.GalvanicICResponse = 1, True, False)
                                    dr("GALVANIC_IC_RESPONSE") = cp.GalvanicICResponse
                                    dr("ID") = cp.ID
                                    dr("INSPECTION_ID") = cp.InspectionID
                                    dr("QUESTION_ID") = cp.QuestionID
                                    tbCPPipeGalIC.Rows.Add(dr)
                                ElseIf cp.TestedByInspector Then
                                    dr = tbCPPipeInspectorTested.NewRow
                                    dr("BLANK") = String.Empty
                                    dr("Yes") = cp.TestedByInspectorResponse
                                    dr("No") = Not cp.TestedByInspectorResponse
                                    dr("ID") = cp.ID
                                    dr("INSPECTION_ID") = cp.InspectionID
                                    dr("QUESTION_ID") = cp.QuestionID
                                    dr("TESTED_BY_INSPECTOR_RESPONSE") = cp.TestedByInspectorResponse
                                    tbCPPipeInspectorTested.Rows.Add(dr)
                                Else
                                    dr = tbCPPipe.NewRow
                                    dr("Line#") = "3.6.3." + cp.LineNumber.ToString
                                    dr("Pipe#") = cp.TankPipeIndex
                                    dr("Fuel Type") = slPipeFuelType.Item(cp.TankPipeID)
                                    dr("Contact Point") = cp.ContactPoint
                                    dr("Local Reference Cell Placement") = cp.LocalReferCellPlacement
                                    dr("Local/On") = cp.LocalOn
                                    dr("Remote/Off") = cp.RemoteOff
                                    dr("Pass") = IIf(cp.PassFailIncon = 1, True, False)
                                    dr("Fail") = IIf(cp.PassFailIncon = 0, True, False)
                                    dr("Incon") = IIf(cp.PassFailIncon = 2, True, False)
                                    dr("ID") = cp.ID
                                    dr("INSPECTION_ID") = cp.InspectionID
                                    dr("QUESTION_ID") = cp.QuestionID
                                    dr("TANK_PIPE_ID") = cp.TankPipeID
                                    dr("TANK_PIPE_ENTITY_ID") = cp.TankPipeEntityID
                                    dr("PIPE_INDEX") = cp.TankPipeIndex
                                    'dr("TANK_DISPENSER") = cp.TankDispenser
                                    'dr("GALVANIC") = cp.Galvanic
                                    'dr("IMPRESSED_CURRENT") = cp.ImpressedCurrent
                                    dr("PASSFAILINCON") = cp.PassFailIncon
                                    dr("CITATION") = checkList.CCAT
                                    'dr("Question") = "Description of Remote Reference Cell Placement"
                                    dr("LINE_NUMBER") = cp.LineNumber
                                    tbCPPipe.Rows.Add(dr)
                                End If
                            ElseIf checkList.CheckListItemNumber = "3.7.6" And checkList.Show Then
                                If cp.RemoteReferCellPlacement Then
                                    dr = tbCPTermRemote.NewRow
                                    dr("Description of Remote Reference Cell Placement") = cp.LocalReferCellPlacement
                                    dr("ID") = cp.ID
                                    dr("INSPECTION_ID") = cp.InspectionID
                                    dr("QUESTION_ID") = cp.QuestionID
                                    tbCPTermRemote.Rows.Add(dr)
                                ElseIf cp.GalvanicIC Then
                                    dr = tbCPTermGalIC.NewRow
                                    dr("Galvanic") = IIf(cp.GalvanicICResponse = 0, True, False)
                                    dr("Impressed Current") = IIf(cp.GalvanicICResponse = 1, True, False)
                                    dr("GALVANIC_IC_RESPONSE") = cp.GalvanicICResponse
                                    dr("ID") = cp.ID
                                    dr("INSPECTION_ID") = cp.InspectionID
                                    dr("QUESTION_ID") = cp.QuestionID
                                    tbCPTermGalIC.Rows.Add(dr)
                                ElseIf cp.TestedByInspector Then
                                    dr = tbCPTermInspectorTested.NewRow
                                    dr("BLANK") = String.Empty
                                    dr("Yes") = cp.TestedByInspectorResponse
                                    dr("No") = Not cp.TestedByInspectorResponse
                                    dr("ID") = cp.ID
                                    dr("INSPECTION_ID") = cp.InspectionID
                                    dr("QUESTION_ID") = cp.QuestionID
                                    dr("TESTED_BY_INSPECTOR_RESPONSE") = cp.TestedByInspectorResponse
                                    tbCPTermInspectorTested.Rows.Add(dr)
                                Else
                                    dr = tbCPTerm.NewRow
                                    dr("Line#") = "3.7.6." + cp.LineNumber.ToString
                                    dr("Term#") = cp.TankPipeIndex
                                    dr("Fuel Type") = slPipeFuelType.Item(cp.TankPipeID)
                                    dr("Contact Point") = cp.ContactPoint
                                    dr("Local Reference Cell Placement") = cp.LocalReferCellPlacement
                                    dr("Local/On") = cp.LocalOn
                                    dr("Remote/Off") = cp.RemoteOff
                                    dr("Pass") = IIf(cp.PassFailIncon = 1, True, False)
                                    dr("Fail") = IIf(cp.PassFailIncon = 0, True, False)
                                    dr("Incon") = IIf(cp.PassFailIncon = 2, True, False)
                                    dr("ID") = cp.ID
                                    dr("INSPECTION_ID") = cp.InspectionID
                                    dr("QUESTION_ID") = cp.QuestionID
                                    dr("TANK_PIPE_ID") = cp.TankPipeID
                                    dr("TANK_PIPE_ENTITY_ID") = cp.TankPipeEntityID
                                    dr("TERM_INDEX") = cp.TankPipeIndex
                                    'dr("TANK_DISPENSER") = cp.TankDispenser
                                    'dr("GALVANIC") = cp.Galvanic
                                    'dr("IMPRESSED_CURRENT") = cp.ImpressedCurrent
                                    dr("PASSFAILINCON") = cp.PassFailIncon
                                    dr("CITATION") = checkList.CCAT
                                    'dr("Question") = "Description of Remote Reference Cell Placement"
                                    dr("LINE_NUMBER") = cp.LineNumber
                                    tbCPTerm.Rows.Add(dr)
                                End If
                            End If
                        End If
                    End If
                Next

                tb.TableName = "CP"
                tb.DefaultView.Sort = "CL_POSITION"
                tbCPRect.TableName = "CPRect"
                tbCPRect.DefaultView.Sort = "ID, QUESTION_ID"
                tbCPTank.TableName = "CPTank"
                tbCPTank.DefaultView.Sort = "TANK_INDEX, LINE_NUMBER"
                tbCPPipe.TableName = "CPPipe"
                tbCPPipe.DefaultView.Sort = "PIPE_INDEX, LINE_NUMBER"
                tbCPTerm.TableName = "CPTerm"
                tbCPTerm.DefaultView.Sort = "TERM_INDEX, LINE_NUMBER"
                tbCPTankRemote.TableName = "CPTankRemote"
                tbCPPipeRemote.TableName = "CPPipeRemote"
                tbCPTermRemote.TableName = "CPTermRemote"
                tbCPTankGalIC.TableName = "CPTankGalvanic"
                tbCPPipeGalIC.TableName = "CPPipeGalvanic"
                tbCPTermGalIC.TableName = "CPTermGalvanic"
                tbCPTankInspectorTested.TableName = "CPTankInspectorTested"
                tbCPPipeInspectorTested.TableName = "CPPipeInspectorTested"
                tbCPTermInspectorTested.TableName = "CPTermInspectorTested"

                ds.Tables.Add(tb)
                ds.Tables.Add(tbCPRect)
                ds.Tables.Add(tbCPTank)
                ds.Tables.Add(tbCPPipe)
                ds.Tables.Add(tbCPTerm)
                ds.Tables.Add(tbCPTankRemote)
                ds.Tables.Add(tbCPPipeRemote)
                ds.Tables.Add(tbCPTermRemote)
                ds.Tables.Add(tbCPTankGalIC)
                ds.Tables.Add(tbCPPipeGalIC)
                ds.Tables.Add(tbCPTermGalIC)
                ds.Tables.Add(tbCPTankInspectorTested)
                ds.Tables.Add(tbCPPipeInspectorTested)
                ds.Tables.Add(tbCPTermInspectorTested)

                dsRel = New DataRelation("ResponseToCPRectifier", ds.Tables("CP").Columns("QUESTION_ID"), ds.Tables("CPRect").Columns("QUESTION_ID"), False)
                ds.Relations.Add(dsRel)

                dsRel = New DataRelation("ResponseToCPReadingsTankInspectorTested", ds.Tables("CP").Columns("QUESTION_ID"), ds.Tables("CPTankInspectorTested").Columns("QUESTION_ID"), False)
                ds.Relations.Add(dsRel)

                dsRel = New DataRelation("ResponseToCPReadingsTankGalvanic", ds.Tables("CP").Columns("QUESTION_ID"), ds.Tables("CPTankGalvanic").Columns("QUESTION_ID"), False)
                ds.Relations.Add(dsRel)

                dsRel = New DataRelation("ResponseToCPReadingsTankRemote", ds.Tables("CP").Columns("QUESTION_ID"), ds.Tables("CPTankRemote").Columns("QUESTION_ID"), False)
                ds.Relations.Add(dsRel)

                dsRel = New DataRelation("ResponseToCPReadingsTank", ds.Tables("CP").Columns("QUESTION_ID"), ds.Tables("CPTank").Columns("QUESTION_ID"), False)
                ds.Relations.Add(dsRel)

                dsRel = New DataRelation("ResponseToCPReadingsPipeInspectorTested", ds.Tables("CP").Columns("QUESTION_ID"), ds.Tables("CPPipeInspectorTested").Columns("QUESTION_ID"), False)
                ds.Relations.Add(dsRel)

                dsRel = New DataRelation("ResponseToCPReadingsPipeGalvanic", ds.Tables("CP").Columns("QUESTION_ID"), ds.Tables("CPPipeGalvanic").Columns("QUESTION_ID"), False)
                ds.Relations.Add(dsRel)

                dsRel = New DataRelation("ResponseToCPReadingsPipeRemote", ds.Tables("CP").Columns("QUESTION_ID"), ds.Tables("CPPipeRemote").Columns("QUESTION_ID"), False)
                ds.Relations.Add(dsRel)

                dsRel = New DataRelation("ResponseToCPReadingsPipe", ds.Tables("CP").Columns("QUESTION_ID"), ds.Tables("CPPipe").Columns("QUESTION_ID"), False)
                ds.Relations.Add(dsRel)

                dsRel = New DataRelation("ResponseToCPReadingsTermInspectorTested", ds.Tables("CP").Columns("QUESTION_ID"), ds.Tables("CPTermInspectorTested").Columns("QUESTION_ID"), False)
                ds.Relations.Add(dsRel)

                dsRel = New DataRelation("ResponseToCPReadingsTermGalvanic", ds.Tables("CP").Columns("QUESTION_ID"), ds.Tables("CPTermGalvanic").Columns("QUESTION_ID"), False)
                ds.Relations.Add(dsRel)

                dsRel = New DataRelation("ResponseToCPReadingsTermRemote", ds.Tables("CP").Columns("QUESTION_ID"), ds.Tables("CPTermRemote").Columns("QUESTION_ID"), False)
                ds.Relations.Add(dsRel)

                dsRel = New DataRelation("ResponseToCPReadingsTerm", ds.Tables("CP").Columns("QUESTION_ID"), ds.Tables("CPTerm").Columns("QUESTION_ID"), False)
                ds.Relations.Add(dsRel)

                'ds.Tables("CP").DefaultView.Sort = "CL_POSITION"
                'ds.Tables("CPRect").DefaultView.Sort = "ID, QUESTION_ID"
                'ds.Tables("CPTank").DefaultView.Sort = "Tank#, QUESTION_ID"
                'ds.Tables("CPPipe").DefaultView.Sort = "Pipe#, QUESTION_ID"
                'ds.Tables("CPTerm").DefaultView.Sort = "Term#, QUESTION_ID"

                If [readOnly] Then
                    Dim col As DataColumn
                    For Each col In ds.Tables("CP").Columns
                        col.ReadOnly = True
                    Next
                    For Each col In ds.Tables("CPRect").Columns
                        col.ReadOnly = True
                    Next
                    For Each col In ds.Tables("CPTank").Columns
                        col.ReadOnly = True
                    Next
                    For Each col In ds.Tables("CPPipe").Columns
                        col.ReadOnly = True
                    Next
                    For Each col In ds.Tables("CPTerm").Columns
                        col.ReadOnly = True
                    Next
                    For Each col In ds.Tables("CPTankRemote").Columns
                        col.ReadOnly = True
                    Next
                    For Each col In ds.Tables("CPPipeRemote").Columns
                        col.ReadOnly = True
                    Next
                    For Each col In ds.Tables("CPTermRemote").Columns
                        col.ReadOnly = True
                    Next
                    For Each col In ds.Tables("CPTankGalvanic").Columns
                        col.ReadOnly = True
                    Next
                    For Each col In ds.Tables("CPPipeGalvanic").Columns
                        col.ReadOnly = True
                    Next
                    For Each col In ds.Tables("CPTermGalvanic").Columns
                        col.ReadOnly = True
                    Next
                    For Each col In ds.Tables("CPTankInspectorTested").Columns
                        col.ReadOnly = True
                    Next
                    For Each col In ds.Tables("CPPipeInspectorTested").Columns
                        col.ReadOnly = True
                    Next
                    For Each col In ds.Tables("CPTermInspectorTested").Columns
                        col.ReadOnly = True
                    Next
                End If

                Return ds
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function TankLeakTable(Optional ByVal [readOnly] As Boolean = False) As DataSet
            Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo
            Dim dsRel As DataRelation
            Dim ds As New DataSet
            Dim dr As DataRow
            Dim tb As New DataTable
            Dim tbWell As New DataTable
            Try
                tb.Columns.Add("CL_POSITION", GetType(Int64))
                tb.Columns.Add("Line#")
                tb.Columns.Add("Question")
                tb.Columns.Add("Yes", GetType(Boolean))
                tb.Columns.Add("No", GetType(Boolean))
                tb.Columns.Add("N/A", GetType(Boolean))
                tb.Columns.Add("CCAT")
                tb.Columns.Add("ID")
                tb.Columns.Add("INSPECTION_ID")
                tb.Columns.Add("QUESTION_ID")
                tb.Columns.Add("SOC")
                tb.Columns.Add("RESPONSE")
                tb.Columns.Add("HEADER")
                tb.Columns.Add("CITATION")
                tb.Columns.Add("FORE_COLOR")
                tb.Columns.Add("BACK_COLOR")

                For Each response As MUSTER.Info.InspectionResponsesInfo In oInspection.ResponsesCollection.Values
                    If Not response.Deleted Then
                        checkList = oInspection.ChecklistMasterCollection.Item(response.QuestionID)
                        If Not checkList Is Nothing Then
                            If checkList.CheckListItemNumber >= "4" And checkList.CheckListItemNumber < "5" And checkList.Show Then
                                dr = tb.NewRow()
                                dr("CL_POSITION") = checkList.Position
                                dr("Line#") = checkList.CheckListItemNumber
                                dr("Question") = checkList.HeaderQuestionText
                                dr("Yes") = IIf(response.Response = 1 Or response.Response = -2, True, False)
                                dr("No") = IIf(response.Response = 0 Or response.Response = -2, True, False)
                                dr("N/A") = IIf(response.Response = 2 Or response.Response = -2, True, False)
                                dr("CCAT") = GetCCAT(checkList)
                                dr("ID") = response.ID
                                dr("INSPECTION_ID") = response.InspectionID
                                dr("QUESTION_ID") = response.QuestionID
                                dr("SOC") = response.SOC
                                dr("RESPONSE") = response.Response
                                dr("HEADER") = checkList.Header
                                dr("CITATION") = checkList.CCAT
                                dr("FORE_COLOR") = checkList.ForeColor
                                dr("BACK_COLOR") = checkList.BackColor
                                tb.Rows.Add(dr)
                            End If
                        End If
                    End If
                Next

                tbWell.Columns.Add("Line#")
                tbWell.Columns.Add("Well#", GetType(Int64))
                tbWell.Columns.Add("Well Depth")
                tbWell.Columns.Add("Depth to" + vbCrLf + "Water")
                tbWell.Columns.Add("Depth to" + vbCrLf + "Slots")
                tbWell.Columns.Add("Surface Sealed" + vbCrLf + "Yes", GetType(Boolean))
                tbWell.Columns.Add("Surface Sealed" + vbCrLf + "No", GetType(Boolean))
                tbWell.Columns.Add("Well Caps" + vbCrLf + "Yes", GetType(Boolean))
                tbWell.Columns.Add("Well Caps" + vbCrLf + "No", GetType(Boolean))
                tbWell.Columns.Add("Inspector's Observations")
                tbWell.Columns.Add("ID")
                tbWell.Columns.Add("INSPECTION_ID")
                tbWell.Columns.Add("QUESTION_ID")
                tbWell.Columns.Add("TANK_LINE")
                tbWell.Columns.Add("SURFACE_SEALED")
                tbWell.Columns.Add("WELL_CAPS")
                tbWell.Columns.Add("CITATION")
                tbWell.Columns.Add("LINE_NUMBER")

                For Each well As MUSTER.Info.InspectionMonitorWellsInfo In oInspection.MonitorWellsCollection.Values
                    If Not well.Deleted Then
                        checkList = oInspection.ChecklistMasterCollection.Item(Math.Abs(IIf(well.QuestionID < -100000, well.QuestionID + 100000, well.QuestionID)))
                        If Not checkList Is Nothing Then
                            If checkList.CheckListItemNumber = "4.2.8" And (checkList.Show Or checkList.ID < -20) Then
                                dr = tbWell.NewRow
                                dr("Line#") = "4.2.8." + well.LineNumber.ToString
                                dr("Well#") = well.WellNumber
                                dr("Well Depth") = well.WellDepth
                                dr("Depth to" + vbCrLf + "Water") = well.DepthToWater
                                dr("Depth to" + vbCrLf + "Slots") = well.DepthToSlots
                                dr("Surface Sealed" + vbCrLf + "Yes") = IIf(well.SurfaceSealed = 1, True, False)
                                dr("Surface Sealed" + vbCrLf + "No") = IIf(well.SurfaceSealed = 0, True, False)
                                dr("Well Caps" + vbCrLf + "Yes") = IIf(well.WellCaps = 1, True, False)
                                dr("Well Caps" + vbCrLf + "No") = IIf(well.WellCaps = 0, True, False)
                                dr("Inspector's Observations") = well.InspectorsObservations
                                dr("ID") = well.ID
                                dr("INSPECTION_ID") = well.InspectionID
                                dr("QUESTION_ID") = well.QuestionID
                                dr("QUESTION_ID") = Math.Abs(IIf(well.QuestionID < -100000, well.QuestionID + 100000, well.QuestionID))
                                dr("TANK_LINE") = well.TankLine
                                dr("SURFACE_SEALED") = well.SurfaceSealed
                                dr("WELL_CAPS") = well.WellCaps
                                dr("CITATION") = checkList.CCAT
                                dr("LINE_NUMBER") = well.LineNumber
                                tbWell.Rows.Add(dr)
                            End If
                        End If
                    End If
                Next

                tb.TableName = "TankLeak"
                tbWell.TableName = "Well"

                ds.Tables.Add(tb)
                ds.Tables.Add(tbWell)

                dsRel = New DataRelation("TankLeakWell", ds.Tables("TankLeak").Columns("QUESTION_ID"), ds.Tables("Well").Columns("QUESTION_ID"), False)
                ds.Relations.Add(dsRel)

                ds.Tables("TankLeak").DefaultView.Sort = "CL_POSITION"
                ds.Tables("Well").DefaultView.Sort = "Well#, LINE_NUMBER"

                If [readOnly] Then
                    Dim col As DataColumn
                    For Each col In ds.Tables("TankLeak").Columns
                        col.ReadOnly = True
                    Next
                    For Each col In ds.Tables("Well").Columns
                        col.ReadOnly = True
                    Next
                End If
                Return ds
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function PipeLeakTable(Optional ByVal [readOnly] As Boolean = False) As DataSet
            Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo
            Dim dsRel As DataRelation
            Dim ds As New DataSet
            Dim dr As DataRow
            Dim tb As New DataTable
            Dim tbWell As New DataTable
            Try
                tb.Columns.Add("CL_POSITION", GetType(Int64))
                tb.Columns.Add("Line#")
                tb.Columns.Add("Question")
                tb.Columns.Add("Yes", GetType(Boolean))
                tb.Columns.Add("No", GetType(Boolean))
                tb.Columns.Add("N/A", GetType(Boolean))
                tb.Columns.Add("CCAT")
                tb.Columns.Add("ID")
                tb.Columns.Add("INSPECTION_ID")
                tb.Columns.Add("QUESTION_ID")
                tb.Columns.Add("SOC")
                tb.Columns.Add("RESPONSE")
                tb.Columns.Add("HEADER")
                tb.Columns.Add("CITATION")
                tb.Columns.Add("FORE_COLOR")
                tb.Columns.Add("BACK_COLOR")

                For Each response As MUSTER.Info.InspectionResponsesInfo In oInspection.ResponsesCollection.Values
                    If Not response.Deleted Then
                        checkList = oInspection.ChecklistMasterCollection.Item(response.QuestionID)
                        If Not checkList Is Nothing Then
                            If checkList.CheckListItemNumber >= "5" And checkList.CheckListItemNumber < "5.9" And checkList.Show Then
                                dr = tb.NewRow()
                                dr("CL_POSITION") = checkList.Position
                                dr("Line#") = checkList.CheckListItemNumber
                                dr("Question") = checkList.HeaderQuestionText
                                dr("Yes") = IIf(response.Response = 1 Or response.Response = -2, True, False)
                                dr("No") = IIf(response.Response = 0 Or response.Response = -2, True, False)
                                dr("N/A") = IIf(response.Response = 2 Or response.Response = -2, True, False)

                                dr("CCAT") = GetCCAT(checkList)
                                dr("ID") = response.ID
                                dr("INSPECTION_ID") = response.InspectionID
                                dr("QUESTION_ID") = response.QuestionID
                                dr("SOC") = response.SOC
                                dr("RESPONSE") = response.Response
                                dr("HEADER") = checkList.Header
                                dr("CITATION") = checkList.CCAT
                                dr("FORE_COLOR") = checkList.ForeColor
                                dr("BACK_COLOR") = checkList.BackColor
                                tb.Rows.Add(dr)
                            End If
                        End If
                    End If
                Next

                tbWell.Columns.Add("Line#")
                tbWell.Columns.Add("Well#", GetType(Int64))
                tbWell.Columns.Add("Well Depth")
                tbWell.Columns.Add("Depth to" + vbCrLf + "Water")
                tbWell.Columns.Add("Depth to" + vbCrLf + "Slots")
                tbWell.Columns.Add("Surface Sealed" + vbCrLf + "Yes", GetType(Boolean))
                tbWell.Columns.Add("Surface Sealed" + vbCrLf + "No", GetType(Boolean))
                tbWell.Columns.Add("Well Caps" + vbCrLf + "Yes", GetType(Boolean))
                tbWell.Columns.Add("Well Caps" + vbCrLf + "No", GetType(Boolean))
                tbWell.Columns.Add("Inspector's Observations")
                tbWell.Columns.Add("ID")
                tbWell.Columns.Add("INSPECTION_ID")
                tbWell.Columns.Add("QUESTION_ID")
                tbWell.Columns.Add("TANK_LINE")
                tbWell.Columns.Add("SURFACE_SEALED")
                tbWell.Columns.Add("WELL_CAPS")
                tbWell.Columns.Add("CITATION")
                tbWell.Columns.Add("LINE_NUMBER")

                For Each well As MUSTER.Info.InspectionMonitorWellsInfo In oInspection.MonitorWellsCollection.Values
                    If Not well.Deleted Then
                        checkList = oInspection.ChecklistMasterCollection.Item(Math.Abs(IIf(well.QuestionID < -100000, well.QuestionID + 100000, well.QuestionID)))
                        If Not checkList Is Nothing Then
                            If checkList.CheckListItemNumber = "5.2.8" AndAlso (checkList.Show Or checkList.ID < -20) Then
                                dr = tbWell.NewRow
                                dr("Line#") = "5.2.8." + well.LineNumber.ToString
                                dr("Well#") = well.WellNumber
                                dr("Well Depth") = well.WellDepth
                                dr("Depth to" + vbCrLf + "Water") = well.DepthToWater
                                dr("Depth to" + vbCrLf + "Slots") = well.DepthToSlots
                                dr("Surface Sealed" + vbCrLf + "Yes") = IIf(well.SurfaceSealed = 1, True, False)
                                dr("Surface Sealed" + vbCrLf + "No") = IIf(well.SurfaceSealed = 0, True, False)
                                dr("Well Caps" + vbCrLf + "Yes") = IIf(well.WellCaps = 1, True, False)
                                dr("Well Caps" + vbCrLf + "No") = IIf(well.WellCaps = 0, True, False)
                                dr("Inspector's Observations") = well.InspectorsObservations
                                dr("ID") = well.ID
                                dr("INSPECTION_ID") = well.InspectionID
                                dr("QUESTION_ID") = Math.Abs(IIf(well.QuestionID < -100000, well.QuestionID + 100000, well.QuestionID))
                                dr("TANK_LINE") = well.TankLine
                                dr("SURFACE_SEALED") = well.SurfaceSealed
                                dr("WELL_CAPS") = well.WellCaps
                                dr("CITATION") = checkList.CCAT
                                dr("LINE_NUMBER") = well.LineNumber
                                tbWell.Rows.Add(dr)
                            End If
                        End If
                    End If
                Next

                tb.TableName = "PipeLeak"
                tbWell.TableName = "Well"

                ds.Tables.Add(tb)
                ds.Tables.Add(tbWell)

                dsRel = New DataRelation("PipeLeakWell", ds.Tables("PipeLeak").Columns("QUESTION_ID"), ds.Tables("Well").Columns("QUESTION_ID"), False)
                ds.Relations.Add(dsRel)

                ds.Tables("PipeLeak").DefaultView.Sort = "CL_POSITION"
                ds.Tables("Well").DefaultView.Sort = "Well#, LINE_NUMBER"

                If [readOnly] Then
                    Dim col As DataColumn
                    For Each col In ds.Tables("PipeLeak").Columns
                        col.ReadOnly = True
                    Next
                    For Each col In ds.Tables("Well").Columns
                        col.ReadOnly = True
                    Next
                End If
                Return ds
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function CATLeakTable(Optional ByVal [readOnly] As Boolean = False) As DataTable
            Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo
            Dim dr As DataRow
            Dim tb As New DataTable
            Try
                tb.Columns.Add("CL_POSITION", GetType(Int64))
                tb.Columns.Add("Line#")
                tb.Columns.Add("Question")
                tb.Columns.Add("Yes", GetType(Boolean))
                tb.Columns.Add("No", GetType(Boolean))
                tb.Columns.Add("N/A", GetType(Boolean))
                tb.Columns.Add("CCAT")
                tb.Columns.Add("ID")
                tb.Columns.Add("INSPECTION_ID")
                tb.Columns.Add("QUESTION_ID")
                tb.Columns.Add("SOC")
                tb.Columns.Add("RESPONSE")
                tb.Columns.Add("HEADER")
                tb.Columns.Add("CITATION")
                tb.Columns.Add("FORE_COLOR")
                tb.Columns.Add("BACK_COLOR")

                For Each response As MUSTER.Info.InspectionResponsesInfo In oInspection.ResponsesCollection.Values
                    If Not response.Deleted Then
                        checkList = oInspection.ChecklistMasterCollection.Item(response.QuestionID)
                        If Not checkList Is Nothing Then
                            If checkList.CheckListItemNumber >= "5.9" And checkList.CheckListItemNumber < "6" And checkList.Show Then
                                dr = tb.NewRow()
                                dr("CL_POSITION") = checkList.Position
                                dr("Line#") = checkList.CheckListItemNumber
                                dr("Question") = checkList.HeaderQuestionText
                                dr("Yes") = IIf(response.Response = 1 Or response.Response = -2, True, False)
                                dr("No") = IIf(response.Response = 0 Or response.Response = -2, True, False)
                                dr("N/A") = IIf(response.Response = 2 Or response.Response = -2, True, False)

                                dr("CCAT") = GetCCAT(checkList)
                                dr("ID") = response.ID
                                dr("INSPECTION_ID") = response.InspectionID
                                dr("QUESTION_ID") = response.QuestionID
                                dr("SOC") = response.SOC
                                dr("RESPONSE") = response.Response
                                dr("HEADER") = checkList.Header
                                dr("CITATION") = checkList.CCAT
                                dr("FORE_COLOR") = checkList.ForeColor
                                dr("BACK_COLOR") = checkList.BackColor
                                tb.Rows.Add(dr)
                            End If
                        End If
                    End If
                Next
                tb.DefaultView.Sort = "CL_POSITION"
                If [readOnly] Then
                    For Each col As DataColumn In tb.Columns
                        col.ReadOnly = True
                    Next
                End If
                Return tb
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function VisualTable(Optional ByVal [readOnly] As Boolean = False) As DataTable
            Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo
            Dim dr As DataRow
            Dim tb As New DataTable
            Try
                tb.Columns.Add("CL_POSITION", GetType(Int64))
                tb.Columns.Add("Line#")
                tb.Columns.Add("Question")
                tb.Columns.Add("Yes", GetType(Boolean))
                tb.Columns.Add("No", GetType(Boolean))
                tb.Columns.Add("N/A", GetType(Boolean))
                tb.Columns.Add("CCAT")
                tb.Columns.Add("ID")
                tb.Columns.Add("INSPECTION_ID")
                tb.Columns.Add("QUESTION_ID")
                tb.Columns.Add("SOC")
                tb.Columns.Add("RESPONSE")
                tb.Columns.Add("HEADER")
                tb.Columns.Add("CITATION")
                tb.Columns.Add("FORE_COLOR")
                tb.Columns.Add("BACK_COLOR")

                For Each response As MUSTER.Info.InspectionResponsesInfo In oInspection.ResponsesCollection.Values
                    If Not response.Deleted Then
                        checkList = oInspection.ChecklistMasterCollection.Item(response.QuestionID)
                        If Not checkList Is Nothing Then
                            If checkList.CheckListItemNumber >= "6" And checkList.CheckListItemNumber < "7" And checkList.Show Then
                                dr = tb.NewRow()
                                dr("CL_POSITION") = checkList.Position
                                dr("Line#") = checkList.CheckListItemNumber
                                dr("Question") = checkList.HeaderQuestionText
                                dr("Yes") = IIf(response.Response = 1 Or response.Response = -2, True, False)
                                dr("No") = IIf(response.Response = 0 Or response.Response = -2, True, False)
                                dr("N/A") = IIf(response.Response = 2 Or response.Response = -2, True, False)

                                dr("CCAT") = GetCCAT(checkList)
                                dr("ID") = response.ID
                                dr("INSPECTION_ID") = response.InspectionID
                                dr("QUESTION_ID") = response.QuestionID
                                dr("SOC") = response.SOC
                                dr("RESPONSE") = response.Response
                                dr("HEADER") = checkList.Header
                                dr("CITATION") = checkList.CCAT
                                dr("FORE_COLOR") = checkList.ForeColor
                                dr("BACK_COLOR") = checkList.BackColor
                                tb.Rows.Add(dr)
                            End If
                        End If
                    End If
                Next
                tb.DefaultView.Sort = "CL_POSITION"
                If [readOnly] Then
                    For Each col As DataColumn In tb.Columns
                        col.ReadOnly = True
                    Next
                End If
                Return tb
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function TOSTable(Optional ByVal [readOnly] As Boolean = False) As DataTable
            Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo
            Dim dr As DataRow
            Dim tb As New DataTable
            Try
                tb.Columns.Add("CL_POSITION", GetType(Int64))
                tb.Columns.Add("Line#")
                tb.Columns.Add("Question")
                tb.Columns.Add("Yes", GetType(Boolean))
                tb.Columns.Add("No", GetType(Boolean))
                tb.Columns.Add("N/A", GetType(Boolean))
                tb.Columns.Add("CCAT")
                tb.Columns.Add("ID")
                tb.Columns.Add("INSPECTION_ID")
                tb.Columns.Add("QUESTION_ID")
                tb.Columns.Add("SOC")
                tb.Columns.Add("RESPONSE")
                tb.Columns.Add("HEADER")
                tb.Columns.Add("CITATION")
                tb.Columns.Add("FORE_COLOR")
                tb.Columns.Add("BACK_COLOR")

                For Each response As MUSTER.Info.InspectionResponsesInfo In oInspection.ResponsesCollection.Values
                    If Not response.Deleted Then
                        checkList = oInspection.ChecklistMasterCollection.Item(response.QuestionID)
                        If Not checkList Is Nothing Then
                            If checkList.CheckListItemNumber >= "7" And checkList.CheckListItemNumber < "8" And checkList.Show Then
                                dr = tb.NewRow()
                                dr("CL_POSITION") = checkList.Position
                                dr("Line#") = checkList.CheckListItemNumber
                                dr("Question") = checkList.HeaderQuestionText
                                dr("Yes") = IIf(response.Response = 1 Or response.Response = -2, True, False)
                                dr("No") = IIf(response.Response = 0 Or response.Response = -2, True, False)
                                dr("N/A") = IIf(response.Response = 2 Or response.Response = -2, True, False)

                                dr("CCAT") = GetCCAT(checkList)
                                dr("ID") = response.ID
                                dr("INSPECTION_ID") = response.InspectionID
                                dr("QUESTION_ID") = response.QuestionID
                                dr("SOC") = response.SOC
                                dr("RESPONSE") = response.Response
                                dr("HEADER") = checkList.Header
                                dr("CITATION") = checkList.CCAT
                                dr("FORE_COLOR") = checkList.ForeColor
                                dr("BACK_COLOR") = checkList.BackColor
                                tb.Rows.Add(dr)
                            End If
                        End If
                    End If
                Next
                tb.DefaultView.Sort = "CL_POSITION"
                If [readOnly] Then
                    For Each col As DataColumn In tb.Columns
                        col.ReadOnly = True
                    Next
                End If
                Return tb
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function CitationTable(Optional ByVal [readOnly] As Boolean = False) As DataTable
            Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo
            Dim dr As DataRow
            Dim tb As New DataTable
            Try
                tb.Columns.Add("CL_POSITION", GetType(Int64))
                tb.Columns.Add("Line#")
                tb.Columns.Add("Citation")
                tb.Columns.Add("CCAT")
                tb.Columns.Add("ID")
                tb.Columns.Add("INSPECTION_ID")
                tb.Columns.Add("QUESTION_ID")
                tb.Columns.Add("DELETED")
                tb.Columns.Add("FORE_COLOR")
                tb.Columns.Add("BACK_COLOR")

                For Each citation As MUSTER.Info.InspectionCitationInfo In oInspection.CitationsCollection.Values
                    checkList = oInspection.ChecklistMasterCollection.Item(citation.QuestionID)
                    If Not checkList Is Nothing Then
                        If Not citation.Deleted And (checkList.Show OrElse checkList.ID < -20) And (checkList.ID <> 56 AndAlso checkList.ID <> 94) Then
                            dr = tb.NewRow()
                            dr("CL_POSITION") = checkList.Position
                            dr("Line#") = checkList.CheckListItemNumber
                            dr("Citation") = checkList.DiscrepText
                            dr("CCAT") = citation.CCAT
                            dr("ID") = citation.ID
                            dr("INSPECTION_ID") = citation.InspectionID
                            dr("QUESTION_ID") = citation.QuestionID
                            dr("DELETED") = citation.Deleted
                            dr("FORE_COLOR") = "BLACK"
                            dr("BACK_COLOR") = "WHITE"
                            tb.Rows.Add(dr)
                        End If
                    End If
                Next
                tb.DefaultView.Sort = "CL_POSITION"
                If [readOnly] Then
                    For Each col As DataColumn In tb.Columns
                        col.ReadOnly = True
                    Next
                End If
                Return tb
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function CCATTankTable(ByVal qid As Int64, Optional ByVal [readOnly] As Boolean = False, Optional ByVal tankID As Integer = 0) As DataTable
            Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo
            Dim tnk As New MUSTER.Info.TankInfo
            Dim comp As New MUSTER.BusinessLogic.pCompartment
            Dim dr As DataRow

            Dim tb As New DataTable
            Dim strSubstance, strFuelType As String
            Dim oProperty As New MUSTER.BusinessLogic.pProperty
            Try
                tb.Columns.Add("Substance")
                tb.Columns.Add("FuelType")
                tb.Columns.Add("Tank#", GetType(Integer))
                tb.Columns.Add("CCAT", GetType(Boolean))
                tb.Columns.Add("CompartmentID", GetType(Integer))
                tb.Columns.Add("Additional Details")
                tb.Columns.Add("ID")
                tb.Columns.Add("INSPECTION_ID")
                tb.Columns.Add("QUESTION_ID")
                tb.Columns.Add("DELETED")

                For Each ccat As MUSTER.Info.InspectionCCATInfo In oInspection.CCATsCollection.Values

                    If Not ccat.Deleted Then


                        If ccat.QuestionID = qid And ccat.TankPipeEntityID = 12 Then

                            tnk = oOwner.Facility.TankCollection.Item(ccat.TankPipeID)


                            checkList = oInspection.ChecklistMasterCollection.Item(ccat.QuestionID)
                            If Not checkList Is Nothing Then

                                strSubstance = String.Empty
                                strFuelType = String.Empty
                                dr = tb.NewRow()

                                comp.Retrieve(tnk, tnk.TankId)

                                If Not comp.CompartmentCollection Is Nothing Then

                                    For Each item As Info.CompartmentInfo In comp.CompartmentCollection.Values

                                        If item.COMPARTMENTNumber = ccat.CompartmentID Then

                                            dr("Substance") = oProperty.GetPropertyNameByID(item.Substance)
                                            dr("FuelType") = oProperty.GetPropertyNameByID(item.FuelTypeId)
                                        End If


                                    Next
                                Else
                                    dr("Substance") = comp.SubstanceDesc
                                    dr("FuelType") = comp.FuelTypeIdDesc

                                End If


                                dr("Tank#") = tnk.TankIndex
                                dr("CCAT") = ccat.TankPipeResponse
                                dr("CompartmentID") = ccat.CompartmentID
                                dr("Additional Details") = ccat.TankPipeResponseDetail
                                dr("ID") = ccat.ID
                                dr("INSPECTION_ID") = ccat.InspectionID
                                dr("QUESTION_ID") = ccat.QuestionID
                                dr("DELETED") = ccat.Deleted

                                tb.Rows.Add(dr)

                            End If

                        End If
                    End If
                Next
                If tb.Rows.Count = 0 And tankID <> 0 Then


                    Dim ccat As New MUSTER.Info.InspectionCCATInfo(0, _
                    oInspection.ID, _
                    qid, _
                    tankID, _
                    12, _
                    False, _
                    False, _
                    String.Empty, _
                    False, _
                    String.Empty, _
                    CDate("01/01/0001"), _
                    String.Empty, _
                    CDate("01/01/0001"))



                    oInspectionCCAT.Add(ccat)



                    'Dim list As New Collections.ArrayList

                    tnk = oOwner.Facility.TankCollection.Item(ccat.TankPipeID)

                    'If Not tnk.CompartmentCollection Is Nothing AndAlso tnk.CompartmentCollection.Count > 0 Then
                    'For Each comp In tnk.CompartmentCollection.Values
                    'list.Add(comp)
                    'Next
                    ' Else
                    '    list.Add(New BusinessLogic.pCompartment)
                    'End If


                    '                For Each item As BusinessLogic.pCompartment In list


                    dr = tb.NewRow()
                    dr("Substance") = String.Empty
                    dr("FuelType") = String.Empty
                    dr("Tank#") = tnk.TankIndex
                    dr("CCAT") = ccat.TankPipeResponse
                    dr("CompartmentID") = ccat.CompartmentID
                    dr("Additional Details") = ccat.TankPipeResponseDetail
                    dr("ID") = ccat.ID
                    dr("INSPECTION_ID") = ccat.InspectionID
                    dr("QUESTION_ID") = ccat.QuestionID
                    dr("DELETED") = ccat.Deleted
                    tb.Rows.Add(dr)

                    '               Next

                End If
                tb.DefaultView.Sort = "Tank#, QUESTION_ID"
                If [readOnly] Then
                    For Each col As DataColumn In tb.Columns
                        col.ReadOnly = True
                    Next
                End If


                comp = Nothing
                tnk = Nothing

                Return tb
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function CCATPipeTable(ByVal qid As Int64, Optional ByVal [readOnly] As Boolean = False, Optional ByVal pipeID As Integer = 0) As DataSet
            Dim tnk As MUSTER.Info.TankInfo
            Dim pipe As MUSTER.BusinessLogic.pPipe
            Dim comp As New MUSTER.BusinessLogic.pCompartment
            Dim strSubstance, strFuelType As String
            Dim oProperty As New MUSTER.BusinessLogic.pProperty

            Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo
            Dim dr As DataRow
            Dim ts As New DataSet
            Dim tb As New DataTable
            Dim tb2 As New DataTable

            Try

                tb.Columns.Add("Substance")
                tb.Columns.Add("FuelType")
                tb.Columns.Add("Pipe#", GetType(Integer))
                tb.Columns.Add("CCAT", GetType(Boolean))
                tb.Columns.Add("Additional Details")
                tb.Columns.Add("ID")
                tb.Columns.Add("INSPECTION_ID")
                tb.Columns.Add("QUESTION_ID")
                tb.Columns.Add("DELETED")
                tb.Columns.Add("PIPEID")

                tb2.Columns.Add("Substance")
                tb2.Columns.Add("FuelType")
                tb2.Columns.Add("Pipe#", GetType(Integer))
                tb2.Columns.Add("CCAT", GetType(Boolean))
                tb2.Columns.Add("Additional Details")
                tb2.Columns.Add("ID")
                tb2.Columns.Add("INSPECTION_ID")
                tb2.Columns.Add("QUESTION_ID")
                tb2.Columns.Add("DELETED")
                tb2.Columns.Add("PARENTPIPE")

                For Each ccat As MUSTER.Info.InspectionCCATInfo In oInspection.CCATsCollection.Values
                    If Not ccat.Deleted Then
                        If ccat.QuestionID = qid And ccat.TankPipeEntityID = 10 And Not ccat.Termination And ccat.TankPipeID = IIf(pipeID = 0, ccat.TankPipeID, pipeID) Then
                            checkList = oInspection.ChecklistMasterCollection.Item(ccat.QuestionID)
                            If Not checkList Is Nothing Then
                                Dim strTankID As String = oOwner.RunSQLQuery("SELECT TANK_ID FROM vINSPECTION_PIPES_DISPLAY_DATA WHERE PIPE_ID = " + ccat.TankPipeID.ToString).Tables(0).Rows(0)(0)
                                tnk = oOwner.Facility.TankCollection.Item(strTankID)
                                strSubstance = String.Empty
                                strFuelType = String.Empty

                                comp.Retrieve(tnk, tnk.TankId)


                                'dr = tb.NewRow()


                                If Not comp.CompartmentCollection Is Nothing Then

                                    For Each item As Info.CompartmentInfo In comp.CompartmentCollection.Values

                                        If IIf(item.COMPARTMENTNumber <= 0, 1, item.COMPARTMENTNumber) = IIf(ccat.CompartmentID <= 0, 1, ccat.CompartmentID) Then

                                            strSubstance = oProperty.GetPropertyNameByID(item.Substance)
                                            strFuelType = oProperty.GetPropertyNameByID(item.FuelTypeId)
                                        End If


                                    Next
                                Else
                                    strSubstance = comp.SubstanceDesc
                                    strFuelType = comp.FuelTypeIdDesc

                                End If
                                ' For Each comp In tnk.CompartmentCollection.Values
                                'strSubstance += IIf(comp.Substance = 0, "N/A", oProperty.GetPropertyNameByID(comp.Substance)) + ", "
                                'strFuelType += IIf(comp.FuelTypeId = 0, "N/A", oProperty.GetPropertyNameByID(comp.FuelTypeId)) + ", "
                                'Next
                                'If strSubstance.Length > 0 Then
                                'strSubstance = strSubstance.Trim.TrimEnd(",")
                                'End If
                                'If strFuelType.Length > 0 Then
                                '   strFuelType = strFuelType.Trim.TrimEnd(",")
                                'End If

                                'dr = tb.NewRow()

                                pipe = New BusinessLogic.pPipe

                                Dim key As String = String.Empty

                                For Each keyStr As String In tnk.pipesCollection.GetKeys()
                                    If keyStr.EndsWith(String.Format("|{0}", ccat.TankPipeID)) Then
                                        key = keyStr
                                    End If
                                Next

                                pipe.Retrieve(tnk, key, comp.CompInfo, False)

                                If pipe.ParentPipeID > 0 Then
                                    dr = tb2.NewRow
                                    dr("PARENTPIPE") = pipe.ParentPipeID
                                Else
                                    dr = tb.NewRow
                                    dr("PIPEID") = pipe.PipeID

                                End If


                                Dim strPipeIndex As String = ""

                                strPipeIndex = pipe.Index.ToString

                                If strPipeIndex = String.Empty Then
                                    dr("Pipe#") = 0
                                Else
                                    dr("Pipe#") = strPipeIndex
                                End If

                                dr("Substance") = strSubstance
                                dr("FuelType") = strFuelType
                                dr("CCAT") = ccat.TankPipeResponse
                                dr("Additional Details") = ccat.TankPipeResponseDetail
                                dr("ID") = ccat.ID
                                dr("INSPECTION_ID") = ccat.InspectionID
                                dr("QUESTION_ID") = ccat.QuestionID
                                dr("DELETED") = ccat.Deleted

                                If pipe.ParentPipeID > 0 Then
                                    tb2.Rows.Add(dr)
                                Else
                                    tb.Rows.Add(dr)
                                End If

                                pipe = Nothing

                            End If
                        End If
                    End If
                Next
                If tb.Rows.Count = 0 And pipeID <> 0 Then
                    Dim ccat As New MUSTER.Info.InspectionCCATInfo(0, _
                    oInspection.ID, _
                    qid, _
                    pipeID, _
                    10, _
                    False, _
                    False, _
                    String.Empty, _
                    False, _
                    String.Empty, _
                    CDate("01/01/0001"), _
                    String.Empty, _
                    CDate("01/01/0001"))
                    oInspectionCCAT.Add(ccat)

                    Dim strTankID As String = oOwner.RunSQLQuery("SELECT TANK_ID FROM vINSPECTION_PIPES_DISPLAY_DATA WHERE PIPE_ID = " + ccat.TankPipeID.ToString).Tables(0).Rows(0)(0)
                    tnk = oOwner.Facility.TankCollection.Item(strTankID)
                    strSubstance = String.Empty
                    strFuelType = String.Empty
                    dr = tb.NewRow()
                    'For Each comp In tnk.CompartmentCollection.Values
                    'strSubstance += IIf(comp.Substance = 0, "N/A", oProperty.GetPropertyNameByID(comp.Substance)) + ", "
                    'strFuelType += IIf(comp.FuelTypeId = 0, "N/A", oProperty.GetPropertyNameByID(comp.FuelTypeId)) + ", "
                    'Next
                    'If strSubstance.Length > 0 Then
                    'strSubstance = strSubstance.Trim.TrimEnd(",")
                    'End If
                    'If strFuelType.Length > 0 Then
                    '   strFuelType = strFuelType.Trim.TrimEnd(",")
                    'End If
                    dr("Substance") = strSubstance
                    dr("FuelType") = strFuelType

                    dr("Pipe#") = slPipeID.Item(ccat.TankPipeID)
                    dr("CCAT") = ccat.TankPipeResponse
                    dr("Additional Details") = ccat.TankPipeResponseDetail
                    dr("ID") = ccat.ID
                    dr("INSPECTION_ID") = ccat.InspectionID
                    dr("QUESTION_ID") = ccat.QuestionID
                    dr("DELETED") = ccat.Deleted
                    dr("PIPEID") = 0
                    tb.Rows.Add(dr)
                End If
                tb.DefaultView.Sort = "Pipe#, QUESTION_ID"
                tb2.DefaultView.Sort = "Pipe#, QUESTION_ID"

                If [readOnly] Then
                    For Each col As DataColumn In tb.Columns
                        col.ReadOnly = True
                    Next
                End If

                ts.Tables.Add(tb)
                ts.Tables.Add(tb2)

                If tb2.Rows.Count > 0 Then

                    Dim List As New Collections.ArrayList
                    Dim ListDict As New Collections.Specialized.ListDictionary
                    Dim tbView As DataTable = tb.Copy




                    For Each drr As DataRow In tb2.Rows
                        List.Add(String.Format("{0};{1}", drr("Pipe#"), drr("PARENTPIPE")))
                        ListDict.Add(String.Format("{0};{1}", drr("Pipe#"), drr("PARENTPIPE")), drr)
                    Next


                    For Each item As String In List
                        If tbView.Select(String.Format("PipeID={0}", item.Substring(item.IndexOf(";") + 1))).GetUpperBound(0) <= -1 Then
                            Dim newR As DataRow = tb.NewRow


                            With DirectCast(ListDict.Item(item), DataRow)
                                newR("Substance") = .Item("Substance")
                                newR("FuelType") = .Item("FuelType")
                                newR("Pipe#") = .Item("Pipe#")
                                newR("CCAT") = .Item("CCAT")
                                newR("Additional Details") = String.Format("{0}{1}", " Unmatched Child Pipe.  ", .Item("Additional Details"))
                                newR("ID") = .Item("ID")
                                newR("INSPECTION_ID") = .Item("INSPECTION_ID")
                                newR("QUESTION_ID") = .Item("QUESTION_ID")
                                newR("DELETED") = .Item("DELETED")
                                newR("PIPEID") = .Item("PARENTPIPE")
                            End With

                            tb.Rows.Add(newR)

                            tb2.Rows.Remove(ListDict.Item(item))

                        End If

                    Next

                    tbView.Dispose()
                    List.Clear()
                    ListDict.Clear()
                    List = Nothing
                    ListDict = Nothing
                End If



                ts.Relations.Add("ParentChild", tb.Columns("PIPEID"), tb2.Columns("PARENTPIPE"))

                Return ts
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function CCATTermTable(ByVal qid As Int64, Optional ByVal [readOnly] As Boolean = False, Optional ByVal pipeID As Integer = 0) As DataSet

            Dim tnk As MUSTER.Info.TankInfo
            Dim pipe As MUSTER.BusinessLogic.pPipe

            Dim comp As New MUSTER.BusinessLogic.pCompartment
            Dim strSubstance, strFuelType As String
            Dim oProperty As New MUSTER.BusinessLogic.pProperty

            Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo
            Dim dr As DataRow
            Dim ts As New DataSet
            Dim tb As New DataTable
            Dim tb2 As New DataTable

            Try
                tb.Columns.Add("Substance")
                tb.Columns.Add("FuelType")
                tb.Columns.Add("Term#", GetType(Integer))
                tb.Columns.Add("CCAT", GetType(Boolean))
                tb.Columns.Add("Additional Details")
                tb.Columns.Add("ID")
                tb.Columns.Add("INSPECTION_ID")
                tb.Columns.Add("QUESTION_ID")
                tb.Columns.Add("DELETED")
                tb.Columns.Add("PIPEID")


                tb2.Columns.Add("Substance")
                tb2.Columns.Add("FuelType")
                tb2.Columns.Add("Term#", GetType(Integer))
                tb2.Columns.Add("CCAT", GetType(Boolean))
                tb2.Columns.Add("Additional Details")
                tb2.Columns.Add("ID")
                tb2.Columns.Add("INSPECTION_ID")
                tb2.Columns.Add("QUESTION_ID")
                tb2.Columns.Add("DELETED")
                tb2.Columns.Add("PARENTPIPE")

                Dim cnt As Integer = 0
                Dim strPipeIndex As String = ""

                For Each ccat As MUSTER.Info.InspectionCCATInfo In oInspection.CCATsCollection.Values
                    If Not ccat.Deleted Then
                        If ccat.QuestionID = qid And ccat.TankPipeEntityID = 10 And ccat.Termination And ccat.TankPipeID = IIf(pipeID = 0, ccat.TankPipeID, pipeID) Then
                            checkList = oInspection.ChecklistMasterCollection.Item(ccat.QuestionID)
                            If Not checkList Is Nothing Then
                                Dim strTankID As String = oOwner.RunSQLQuery("SELECT TANK_ID FROM vINSPECTION_TERMINATIONS_DISPLAY_DATA WHERE PIPE_ID = " + ccat.TankPipeID.ToString).Tables(0).Rows(0)(0)

                                tnk = oOwner.Facility.TankCollection.Item(strTankID)
                                strSubstance = String.Empty
                                strFuelType = String.Empty
                                If Not tnk Is Nothing Then
                                    comp.Retrieve(tnk, tnk.TankId)

                                    'dr = tb.NewRow()


                                    If Not comp.CompartmentCollection Is Nothing Then

                                        For Each item As Info.CompartmentInfo In comp.CompartmentCollection.Values

                                            If IIf(item.COMPARTMENTNumber <= 0, 1, item.COMPARTMENTNumber) = IIf(ccat.CompartmentID <= 0, 1, ccat.CompartmentID) Then

                                                strSubstance = oProperty.GetPropertyNameByID(item.Substance)
                                                strFuelType = oProperty.GetPropertyNameByID(item.FuelTypeId)
                                            End If


                                        Next
                                    Else
                                        strSubstance = comp.SubstanceDesc
                                        strFuelType = comp.FuelTypeIdDesc

                                    End If
                                    ' For Each comp In tnk.CompartmentCollection.Values
                                    'strSubstance += IIf(comp.Substance = 0, "N/A", oProperty.GetPropertyNameByID(comp.Substance)) + ", "
                                    'strFuelType += IIf(comp.FuelTypeId = 0, "N/A", oProperty.GetPropertyNameByID(comp.FuelTypeId)) + ", "
                                    'Next
                                    'If strSubstance.Length > 0 Then
                                    'strSubstance = strSubstance.Trim.TrimEnd(",")
                                    'End If
                                    'If strFuelType.Length > 0 Then
                                    '   strFuelType = strFuelType.Trim.TrimEnd(",")
                                    'End If

                                    'dr = tb.NewRow()

                                    pipe = New BusinessLogic.pPipe

                                    Dim key As String = String.Empty

                                    For Each keyStr As String In tnk.pipesCollection.GetKeys()
                                        If keyStr.EndsWith(String.Format("|{0}", ccat.TankPipeID)) Then
                                            key = keyStr
                                        End If
                                    Next

                                    pipe.Retrieve(tnk, key, comp.CompInfo, False)

                                    If pipe.ParentPipeID > 0 Then
                                        dr = tb2.NewRow
                                        dr("PARENTPIPE") = pipe.ParentPipeID
                                    Else
                                        dr = tb.NewRow
                                        dr("PIPEID") = pipe.PipeID

                                    End If




                                    strPipeIndex = pipe.Index.ToString

                                    If strPipeIndex = String.Empty Then
                                        dr("term#") = 0
                                    Else
                                        dr("term#") = strPipeIndex
                                    End If

                                    dr("Substance") = strSubstance
                                    dr("FuelType") = strFuelType
                                    dr("CCAT") = ccat.TankPipeResponse
                                    dr("Additional Details") = ccat.TankPipeResponseDetail
                                    dr("ID") = ccat.ID
                                    dr("INSPECTION_ID") = ccat.InspectionID
                                    dr("QUESTION_ID") = ccat.QuestionID
                                    dr("DELETED") = ccat.Deleted

                                    If pipe.ParentPipeID > 0 Then
                                        tb2.Rows.Add(dr)
                                    Else
                                        tb.Rows.Add(dr)
                                    End If
                                End If
                                pipe = Nothing

                                cnt += 1

                            End If
                        End If
                    End If
                Next
                If tb.Rows.Count = 0 And pipeID <> 0 Then
                    Dim ccat As New MUSTER.Info.InspectionCCATInfo(0, _
                    oInspection.ID, _
                    qid, _
                    pipeID, _
                    10, _
                    False, _
                    True, _
                    String.Empty, _
                    False, _
                    String.Empty, _
                    CDate("01/01/0001"), _
                    String.Empty, _
                    CDate("01/01/0001"))
                    oInspectionCCAT.Add(ccat)

                    Dim strTankID As String = oOwner.RunSQLQuery("SELECT TANK_ID FROM vINSPECTION_TERMINATIONS_DISPLAY_DATA WHERE PIPE_ID = " + ccat.TankPipeID.ToString).Tables(0).Rows(0)(0)
                    tnk = oOwner.Facility.TankCollection.Item(strTankID)
                    strSubstance = String.Empty
                    strFuelType = String.Empty
                    dr = tb.NewRow()
                    ' For Each comp In tnk.CompartmentCollection.Values
                    'strSubstance += IIf(comp.Substance = 0, "N/A", oProperty.GetPropertyNameByID(comp.Substance)) + ", "
                    'strFuelType += IIf(comp.FuelTypeId = 0, "N/A", oProperty.GetPropertyNameByID(comp.FuelTypeId)) + ", "
                    'Next
                    'If strSubstance.Length > 0 Then
                    'strSubstance = strSubstance.Trim.TrimEnd(",")
                    'End If
                    'If strFuelType.Length > 0 Then
                    '   strFuelType = strFuelType.Trim.TrimEnd(",")
                    'End If
                    dr("Substance") = strSubstance

                    dr("FuelType") = strFuelType

                    dr("Term#") = slTermPipeID.Item(ccat.TankPipeID)
                    dr("CCAT") = ccat.TankPipeResponse
                    dr("Additional Details") = ccat.TankPipeResponseDetail
                    dr("ID") = ccat.ID
                    dr("INSPECTION_ID") = ccat.InspectionID
                    dr("QUESTION_ID") = ccat.QuestionID
                    dr("DELETED") = ccat.Deleted
                    tb.Rows.Add(dr)
                End If
                tb.DefaultView.Sort = "Term#, QUESTION_ID"
                tb2.DefaultView.Sort = "Term#, QUESTION_ID"

                If [readOnly] Then
                    For Each col As DataColumn In tb.Columns
                        col.ReadOnly = True
                    Next
                End If
                ts.Tables.Add(tb)
                ts.Tables.Add(tb2)


                If tb2.Rows.Count > 0 Then

                    Dim List As New Collections.ArrayList
                    Dim ListDict As New Collections.Specialized.ListDictionary
                    Dim tbView As DataTable = tb.Copy




                    For Each drr As DataRow In tb2.Rows
                        List.Add(String.Format("{0};{1}", drr("term#"), drr("PARENTPIPE")))
                        ListDict.Add(String.Format("{0};{1}", drr("term#"), drr("PARENTPIPE")), drr)
                    Next


                    For Each item As String In List
                        If tbView.Select(String.Format("PipeID={0}", item.Substring(item.IndexOf(";") + 1))).GetUpperBound(0) <= -1 Then
                            Dim newR As DataRow = tb.NewRow


                            With DirectCast(ListDict.Item(item), DataRow)
                                newR("Substance") = .Item("Substance")
                                newR("FuelType") = .Item("FuelType")
                                newR("Term#") = .Item("Term#")
                                newR("CCAT") = .Item("CCAT")
                                newR("Additional Details") = String.Format("{0}{1}", " Unmatched Child Pipe.  ", .Item("Additional Details"))
                                newR("ID") = .Item("ID")
                                newR("INSPECTION_ID") = .Item("INSPECTION_ID")
                                newR("QUESTION_ID") = .Item("QUESTION_ID")
                                newR("DELETED") = .Item("DELETED")
                                newR("PIPEID") = .Item("PARENTPIPE")
                            End With

                            tb.Rows.Add(newR)

                            tb2.Rows.Remove(ListDict.Item(item))

                        End If

                    Next

                    tbView.Dispose()
                    List.Clear()
                    ListDict.Clear()
                    List = Nothing
                    ListDict = Nothing
                End If






                Return ts
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function DiscrepTable(Optional ByVal [readOnly] As Boolean = False) As DataTable
            Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo
            Dim dr As DataRow
            Dim tb As New DataTable
            Try
                tb.Columns.Add("CL_POSITION", GetType(Int64))
                tb.Columns.Add("Line#")
                tb.Columns.Add("Description")
                tb.Columns.Add("ID")
                tb.Columns.Add("INSPECTION_ID")
                tb.Columns.Add("QUESTION_ID")
                tb.Columns.Add("DELETED")
                tb.Columns.Add("FORE_COLOR")
                tb.Columns.Add("BACK_COLOR")

                For Each discrep As MUSTER.Info.InspectionDiscrepInfo In oInspection.DiscrepsCollection.Values
                    checkList = oInspection.ChecklistMasterCollection.Item(discrep.QuestionID)
                    If Not checkList Is Nothing Then
                        If Not discrep.Deleted And checkList.Show Then
                            dr = tb.NewRow()
                            dr("CL_POSITION") = checkList.Position
                            dr("Line#") = checkList.CheckListItemNumber
                            dr("Description") = discrep.Description
                            dr("ID") = discrep.ID
                            dr("INSPECTION_ID") = discrep.InspectionID
                            dr("QUESTION_ID") = discrep.QuestionID
                            dr("DELETED") = discrep.Deleted
                            dr("FORE_COLOR") = "BLACK"
                            dr("BACK_COLOR") = "WHITE"
                            tb.Rows.Add(dr)
                        End If
                    End If
                Next
                tb.DefaultView.Sort = "CL_POSITION"
                If [readOnly] Then
                    For Each col As DataColumn In tb.Columns
                        col.ReadOnly = True
                    Next
                End If
                Return tb
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function SOCTable(Optional ByVal [readOnly] As Boolean = False) As DataTable
            Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo
            Dim dr As DataRow
            Dim tb As New DataTable
            Dim rowIndex, cnt As Integer
            Dim caeAdminOverride As Boolean = False
            Dim citationText As String
            Dim lineNums As String
            Dim bolYes, bolNo As Boolean
            Dim alLPLD As New ArrayList
            Try
                tb.Columns.Add("CL_POSITION", GetType(Int64))
                tb.Columns.Add("Question")
                tb.Columns.Add("Yes", GetType(Boolean))
                tb.Columns.Add("No", GetType(Boolean))
                tb.Columns.Add("Line Numbers")
                tb.Columns.Add("Citations")
                tb.Columns.Add("ID")
                tb.Columns.Add("INSPECTION_ID")
                tb.Columns.Add("FSOC_LK_PREVENT")
                tb.Columns.Add("FSOC_LK_PRE_CITATION")
                tb.Columns.Add("FAC_SOC_LK_DETECTION")
                tb.Columns.Add("FSOC_LK_DET_CITATION")
                tb.Columns.Add("FSOC_LK_PRE_LK_DET")
                tb.Columns.Add("DELETED")

                dr = tb.NewRow
                'dr("Line#") = "10.1"
                tb.Rows.Add(dr)
                dr = tb.NewRow
                'dr("Line#") = "10.2"
                tb.Rows.Add(dr)
                dr = tb.NewRow
                'dr("Line#") = "10.3"
                tb.Rows.Add(dr)

                cnt = 0

                For Each checkList In oInspection.ChecklistMasterCollection.Values
                    rowIndex = -1
                    If checkList.CheckListItemNumber = "10.1" Then
                        rowIndex = 0
                        citationText = oInspectionSOC.LeakPreventionCitation
                        lineNums = oInspectionSOC.LeakPreventionLineNumbers
                        bolYes = IIf(oInspectionSOC.LeakPrevention = 1, True, False)
                        bolNo = IIf(oInspectionSOC.LeakPrevention = 0, True, False)
                        cnt += 1
                    ElseIf checkList.CheckListItemNumber = "10.2" Then
                        rowIndex = 1
                        citationText = oInspectionSOC.LeakDetectionCitation
                        lineNums = oInspectionSOC.LeakDetectionLineNumbers
                        bolYes = IIf(oInspectionSOC.LeakDetection = 1, True, False)
                        bolNo = IIf(oInspectionSOC.LeakDetection = 0, True, False)
                        cnt += 1
                    ElseIf checkList.CheckListItemNumber = "10.3" Then
                        rowIndex = 2
                        citationText = String.Empty
                        bolYes = IIf(oInspectionSOC.LeakPreventionDetection = 1, True, False)
                        bolNo = IIf(oInspectionSOC.LeakPreventionDetection = 0, True, False)
                        cnt += 1
                    End If
                    If rowIndex <> -1 Then
                        dr = tb.Rows(rowIndex)
                        dr("CL_POSITION") = checkList.Position
                        dr("Question") = checkList.HeaderQuestionText
                        dr("Line Numbers") = lineNums
                        dr("Citations") = citationText
                        dr("Yes") = bolYes
                        dr("No") = bolNo
                        dr("ID") = oInspectionSOC.ID
                        dr("INSPECTION_ID") = oInspection.ID
                        dr("FSOC_LK_PREVENT") = oInspectionSOC.LeakPrevention
                        dr("FSOC_LK_PRE_CITATION") = oInspectionSOC.LeakPreventionCitation
                        dr("FAC_SOC_LK_DETECTION") = oInspectionSOC.LeakDetection
                        dr("FSOC_LK_DET_CITATION") = oInspectionSOC.LeakDetectionCitation
                        dr("FSOC_LK_PRE_LK_DET") = oInspectionSOC.LeakPreventionDetection
                        dr("DELETED") = oInspectionSOC.Deleted
                        If cnt > 2 Then Exit For
                    End If
                Next

                If (oInspectionSOC.ID <= 0 Or Not oInspectionSOC.CAEOverride) And Not [readOnly] Then
                    ' calculate values from response collection
                    ' reset values in table
                    tb.Rows(0)("Yes") = True
                    tb.Rows(0)("No") = False
                    tb.Rows(0)("Line Numbers") = String.Empty
                    tb.Rows(0)("Citations") = String.Empty
                    tb.Rows(1)("Yes") = True
                    tb.Rows(1)("No") = False
                    tb.Rows(1)("Line Numbers") = String.Empty
                    tb.Rows(1)("Citations") = String.Empty
                    tb.Rows(2)("Yes") = True
                    tb.Rows(2)("No") = False
                    tb.Rows(2)("Line Numbers") = String.Empty
                    tb.Rows(2)("Citations") = String.Empty
                    For Each response As MUSTER.Info.InspectionResponsesInfo In oInspection.ResponsesCollection.Values
                        checkList = oInspection.ChecklistMasterCollection.Item(response.QuestionID)
                        If Not checkList Is Nothing Then
                            If (response.SOC And checkList.SOC.Trim <> String.Empty) And response.Response = 0 Then
                                Select Case checkList.SOC.Substring(0, 1).ToUpper
                                    Case "P"
                                        rowIndex = 0
                                    Case "D"
                                        rowIndex = 1
                                End Select
                                If Not alLPLD.Contains(checkList.SOC) Then
                                    alLPLD.Add(checkList.SOC)
                                    tb.Rows(rowIndex)("Citations") += checkList.SOC + ", "
                                End If
                                tb.Rows(rowIndex)("Line Numbers") += checkList.CheckListItemNumber + ", "
                                tb.Rows(rowIndex)("Yes") = False
                                tb.Rows(rowIndex)("No") = True
                            End If
                        End If
                    Next
                    If tb.Rows.Count > 2 Then
                        If tb.Rows(0)("Citations") <> String.Empty Then
                            tb.Rows(0)("Citations") = tb.Rows(0)("Citations").ToString.TrimEnd.TrimEnd(",")
                        End If
                        If tb.Rows(1)("Citations") <> String.Empty Then
                            tb.Rows(1)("Citations") = tb.Rows(1)("Citations").ToString.TrimEnd.TrimEnd(",")
                        End If
                        If tb.Rows(0)("Line Numbers") <> String.Empty Then
                            tb.Rows(0)("Line Numbers") = tb.Rows(0)("Line Numbers").ToString.TrimEnd.TrimEnd(",")
                        End If
                        If tb.Rows(1)("Line Numbers") <> String.Empty Then
                            tb.Rows(1)("Line Numbers") = tb.Rows(1)("Line Numbers").ToString.TrimEnd.TrimEnd(",")
                        End If
                        If tb.Rows(0)("No") Or tb.Rows(1)("No") Then
                            tb.Rows(2)("Yes") = False
                            tb.Rows(2)("No") = True
                        End If
                    End If
                    'Else
                    ' do not calculate values from response collection
                End If
                tb.DefaultView.Sort = "CL_POSITION"
                If [readOnly] Or Date.Compare(oInspection.Completed, CDate("01/01/0001")) <> 0 Then
                    For Each col As DataColumn In tb.Columns
                        col.ReadOnly = True
                    Next
                End If
                Return tb
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function MWellTable(Optional ByVal [readOnly] As Boolean = False) As DataSet
            Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo
            Dim dsRel As DataRelation
            Dim ds As New DataSet
            Dim dr As DataRow
            Dim tb As New DataTable
            Dim tbWell As New DataTable
            Dim bolQuestionIsVisible As Boolean = False
            Try
                tb.Columns.Add("CL_POSITION", GetType(Int64))
                tb.Columns.Add("Line#")
                tb.Columns.Add("Question")
                tb.Columns.Add("Yes", GetType(Boolean))
                tb.Columns.Add("No", GetType(Boolean))
                tb.Columns.Add("N/A", GetType(Boolean))
                tb.Columns.Add("CCAT")
                tb.Columns.Add("ID")
                tb.Columns.Add("INSPECTION_ID")
                tb.Columns.Add("QUESTION_ID")
                tb.Columns.Add("SOC")
                tb.Columns.Add("RESPONSE")
                tb.Columns.Add("HEADER")
                tb.Columns.Add("CITATION")
                tb.Columns.Add("FORE_COLOR")
                tb.Columns.Add("BACK_COLOR")

                For Each response As MUSTER.Info.InspectionResponsesInfo In oInspection.ResponsesCollection.Values
                    If Not response.Deleted Then
                        checkList = oInspection.ChecklistMasterCollection.Item(response.QuestionID)
                        If Not checkList Is Nothing Then
                            If checkList.CheckListItemNumber = "11" Then
                                If checkList.Show Then
                                    dr = tb.NewRow()
                                    dr("CL_POSITION") = checkList.Position
                                    dr("Line#") = checkList.CheckListItemNumber
                                    dr("Question") = checkList.HeaderQuestionText
                                    dr("Yes") = IIf(response.Response = 1, True, False)
                                    dr("No") = IIf(response.Response = 0, True, False)
                                    dr("N/A") = IIf(response.Response = 2, True, False)

                                    dr("CCAT") = GetCCAT(checkList)
                                    dr("ID") = response.ID
                                    dr("INSPECTION_ID") = response.InspectionID
                                    dr("QUESTION_ID") = response.QuestionID
                                    dr("SOC") = response.SOC
                                    dr("RESPONSE") = response.Response
                                    dr("HEADER") = checkList.Header
                                    dr("CITATION") = checkList.CCAT
                                    dr("FORE_COLOR") = checkList.ForeColor
                                    dr("BACK_COLOR") = checkList.BackColor
                                    tb.Rows.Add(dr)
                                    bolQuestionIsVisible = True
                                Else
                                    bolQuestionIsVisible = False
                                End If
                                Exit For
                            End If
                        End If
                    End If
                Next

                tbWell.Columns.Add("Line#")
                tbWell.Columns.Add("Well#", GetType(Int64))
                tbWell.Columns.Add("Well Depth")
                tbWell.Columns.Add("Depth to" + vbCrLf + "Water")
                tbWell.Columns.Add("Depth to" + vbCrLf + "Slots")
                tbWell.Columns.Add("Surface Sealed" + vbCrLf + "Yes", GetType(Boolean))
                tbWell.Columns.Add("Surface Sealed" + vbCrLf + "No", GetType(Boolean))
                tbWell.Columns.Add("Well Caps" + vbCrLf + "Yes", GetType(Boolean))
                tbWell.Columns.Add("Well Caps" + vbCrLf + "No", GetType(Boolean))
                tbWell.Columns.Add("Inspector's Observations")
                tbWell.Columns.Add("ID")
                tbWell.Columns.Add("INSPECTION_ID")
                tbWell.Columns.Add("QUESTION_ID")
                tbWell.Columns.Add("TANK_LINE")
                tbWell.Columns.Add("SURFACE_SEALED")
                tbWell.Columns.Add("WELL_CAPS")
                tbWell.Columns.Add("CITATION")
                tbWell.Columns.Add("LINE_NUMBER")

                If bolQuestionIsVisible Then
                    For Each well As MUSTER.Info.InspectionMonitorWellsInfo In oInspection.MonitorWellsCollection.Values
                        If Not well.Deleted And well.QuestionID = tb.Rows(0)("QUESTION_ID") Then
                            checkList = oInspection.ChecklistMasterCollection.Item(well.QuestionID)
                            If Not checkList Is Nothing Then
                                dr = tbWell.NewRow
                                dr("Line#") = "11." + well.LineNumber.ToString
                                dr("Well#") = well.WellNumber
                                dr("Well Depth") = well.WellDepth
                                dr("Depth to" + vbCrLf + "Water") = well.DepthToWater
                                dr("Depth to" + vbCrLf + "Slots") = well.DepthToSlots
                                dr("Surface Sealed" + vbCrLf + "Yes") = IIf(well.SurfaceSealed = 1, True, False)
                                dr("Surface Sealed" + vbCrLf + "No") = IIf(well.SurfaceSealed = 0, True, False)
                                dr("Well Caps" + vbCrLf + "Yes") = IIf(well.WellCaps = 1, True, False)
                                dr("Well Caps" + vbCrLf + "No") = IIf(well.WellCaps = 0, True, False)
                                dr("Inspector's Observations") = well.InspectorsObservations
                                dr("ID") = well.ID
                                dr("INSPECTION_ID") = well.InspectionID
                                dr("QUESTION_ID") = well.QuestionID
                                dr("TANK_LINE") = well.TankLine
                                dr("SURFACE_SEALED") = well.SurfaceSealed
                                dr("WELL_CAPS") = well.WellCaps
                                dr("CITATION") = checkList.CCAT
                                dr("LINE_NUMBER") = well.LineNumber
                                tbWell.Rows.Add(dr)
                            End If
                        End If
                    Next
                End If

                tb.TableName = "TankPipeMW"
                tbWell.TableName = "Well"

                ds.Tables.Add(tb)
                ds.Tables.Add(tbWell)

                dsRel = New DataRelation("TankPipeMWell", ds.Tables("TankPipeMW").Columns("QUESTION_ID"), ds.Tables("Well").Columns("QUESTION_ID"), False)
                ds.Relations.Add(dsRel)

                ds.Tables("TankPipeMW").DefaultView.Sort = "CL_POSITION"
                ds.Tables("Well").DefaultView.Sort = "Well#, LINE_NUMBER"

                If [readOnly] Then
                    Dim col As DataColumn
                    For Each col In ds.Tables("TankPipeMW").Columns
                        col.ReadOnly = True
                    Next
                    For Each col In ds.Tables("Well").Columns
                        col.ReadOnly = True
                    Next
                End If
                Return ds
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region
#End Region
#Region "External Event Handlers"
        Private Sub InspectionChecklistMasterInfoChanged(ByVal bolValue As Boolean) Handles oInspectionChecklistMasterInfo.evtInspectionChecklistMasterInfoChanged
            RaiseEvent evtInspectionChecklistMasterChanged(bolValue)
        End Sub
        Private Sub InspectionCCATChanged(ByVal bolValue As Boolean) Handles oInspectionCCAT.evtInspectionCCATChanged
            RaiseEvent evtInspectionChecklistMasterChanged(bolValue)
        End Sub
        Private Sub InspectionCitationChanged(ByVal bolValue As Boolean) Handles oInspectionCitation.evtInspectionCitationChanged
            RaiseEvent evtInspectionChecklistMasterChanged(bolValue)
        End Sub
        Private Sub InspectionCommentsChanged(ByVal bolValue As Boolean) Handles oInspectionComments.evtInspectionCommentsChanged
            RaiseEvent evtInspectionChecklistMasterChanged(bolValue)
        End Sub
        Private Sub InspectionCPReadingsChanged(ByVal bolValue As Boolean) Handles oInspectionCPReadings.evtInspectionCPReadingsChanged
            RaiseEvent evtInspectionChecklistMasterChanged(bolValue)
        End Sub
        Private Sub InspectionDescepChanged(ByVal bolValue As Boolean) Handles oInspectionDiscrep.evtInspectionDescepChanged
            RaiseEvent evtInspectionChecklistMasterChanged(bolValue)
        End Sub
        Private Sub InspectionMonitorWellsChanged(ByVal bolValue As Boolean) Handles oInspectionMonitorWells.evtInspectionMonitorWellsChanged
            RaiseEvent evtInspectionChecklistMasterChanged(bolValue)
        End Sub
        Private Sub InspectionRectifierChanged(ByVal bolValue As Boolean) Handles oInspectionRectifier.evtInspectionRectifierChanged
            RaiseEvent evtInspectionChecklistMasterChanged(bolValue)
        End Sub
        Private Sub InspectionResponseChanged(ByVal bolValue As Boolean) Handles oInspectionResponses.evtInspectionResponseChanged
            RaiseEvent evtInspectionChecklistMasterChanged(bolValue)
        End Sub
        Private Sub InspectionSketchChanged(ByVal bolValue As Boolean) Handles oInspectionSketch.evtInspectionSketchChanged
            RaiseEvent evtInspectionChecklistMasterChanged(bolValue)
        End Sub
        Private Sub InspectionSOCChanged(ByVal bolValue As Boolean) Handles oInspectionSOC.evtInspectionSOCChanged
            RaiseEvent evtInspectionChecklistMasterChanged(bolValue)
        End Sub
        Private Sub OwnerChanged(ByVal bolValue As Boolean) Handles oOwner.evtOwnerChanged
            RaiseEvent evtInspectionChecklistMasterChanged(bolValue)
        End Sub
        Private Sub FacilityChanged(ByVal bolValue As Boolean) Handles oOwner.evtFacilityChanged
            RaiseEvent evtInspectionChecklistMasterChanged(bolValue)
        End Sub
        Private Sub AddressChanged(ByVal bolValue As Boolean) Handles oOwner.evtAddressChanged
            RaiseEvent evtInspectionChecklistMasterChanged(bolValue)
        End Sub
        Private Sub AddressesChanged(ByVal bolValue As Boolean) Handles oOwner.evtAddressesChanged
            RaiseEvent evtInspectionChecklistMasterChanged(bolValue)
        End Sub
#End Region
    End Class
End Namespace
