'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.OwnerComplianceEvent
'   Provides the operations required to manipulate an Entity object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        MR        7/1/2005       Original class definition
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
' NOTE: This file to be used as OwnerComplianceEvent to build other objects.
'       Replace keyword "OwnerComplianceEvent" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pOwnerComplianceEvent
#Region "Public Events"
        Public Event OwnerComplianceEventErr(ByVal MsgStr As String)
        Public Event OwnerComplianceEventChanged(ByVal bolValue As Boolean)
        Public Event ColChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        Private WithEvents oOCEInfo As MUSTER.Info.OwnerComplianceEventInfo
        Private WithEvents colOCE As MUSTER.Info.OwnerComplianceEventsCollection
        Private oOCEDB As MUSTER.DataAccess.OwnerComplianceEventDB
        Private MusterException As MUSTER.Exceptions.MusterExceptions
        Private nID As Int64 = -1
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty

        'Private WithEvents pOwn As New MUSTER.BusinessLogic.pOwner
        'Private WithEvents pFacility As New MUSTER.BusinessLogic.pFacility
        'Private pInsCitation As New MUSTER.BusinessLogic.pInspectionCitation
        'Private CitationPenalty As New MUSTER.BusinessLogic.pCitationPenalty
        'Dim dsComplianceDetails As New DataSet
#End Region
#Region "Constructors"
        Public Sub New(Optional ByVal strDBConn As String = "", Optional ByRef MusterXCEP As MUSTER.Exceptions.MusterExceptions = Nothing)
            If MusterXCEP Is Nothing Then
                MusterException = New MUSTER.Exceptions.MusterExceptions
            Else
                MusterException = MusterXCEP
            End If
            oOCEInfo = New MUSTER.Info.OwnerComplianceEventInfo
            colOCE = New MUSTER.Info.OwnerComplianceEventsCollection
            oOCEDB = New MUSTER.DataAccess.OwnerComplianceEventDB
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property OCECollection() As MUSTER.Info.OwnerComplianceEventsCollection
            Get
                Return colOCE
            End Get
            Set(ByVal Value As MUSTER.Info.OwnerComplianceEventsCollection)
                colOCE = Value
            End Set
        End Property
        Public Property OCEInfo() As MUSTER.Info.OwnerComplianceEventInfo
            Get
                Return oOCEInfo
            End Get
            Set(ByVal Value As MUSTER.Info.OwnerComplianceEventInfo)
                oOCEInfo = Value
            End Set
        End Property

        Public Property ID() As Int32
            Get
                Return oOCEInfo.ID
            End Get
            Set(ByVal Value As Int32)
                oOCEInfo.ID = Value
            End Set
        End Property
        Public Property OwnerID() As Int32
            Get
                Return oOCEInfo.OwnerID
            End Get
            Set(ByVal Value As Int32)
                oOCEInfo.OwnerID = Value
            End Set
        End Property
        Public Property Citation() As Integer
            Get
                Return oOCEInfo.Citation
            End Get
            Set(ByVal Value As Integer)
                oOCEInfo.Citation = Value
            End Set
        End Property
        Public Property CitationDueDate() As Date
            Get
                Return oOCEInfo.CitationDueDate
            End Get
            Set(ByVal Value As Date)
                oOCEInfo.CitationDueDate = Value
            End Set
        End Property
        Public Property Rescinded() As Boolean
            Get
                Return oOCEInfo.Rescinded
            End Get
            Set(ByVal Value As Boolean)
                oOCEInfo.Rescinded = Value
            End Set
        End Property
        Public Property OCEPath() As Integer
            Get
                Return oOCEInfo.OCEPath
            End Get
            Set(ByVal Value As Integer)
                oOCEInfo.OCEPath = Value
            End Set
        End Property
        Public Property OCEDate() As Date
            Get
                Return oOCEInfo.OCEDate
            End Get
            Set(ByVal Value As Date)
                oOCEInfo.OCEDate = Value
            End Set
        End Property
        Public Property OCEProcessDate() As Date
            Get
                Return oOCEInfo.OCEProcessDate
            End Get
            Set(ByVal Value As Date)
                oOCEInfo.OCEProcessDate = Value
            End Set
        End Property
        Public Property NextDueDate() As Date
            Get
                Return oOCEInfo.NextDueDate
            End Get
            Set(ByVal Value As Date)
                oOCEInfo.NextDueDate = Value
            End Set
        End Property
        Public Property OverrideDueDate() As Date
            Get
                Return oOCEInfo.OverrideDueDate
            End Get
            Set(ByVal Value As Date)
                oOCEInfo.OverrideDueDate = Value
            End Set
        End Property
        Public Property OCEStatus() As Integer
            Get
                Return oOCEInfo.OCEStatus
            End Get
            Set(ByVal Value As Integer)
                oOCEInfo.OCEStatus = Value
            End Set
        End Property
        Public Property Escalation() As Integer
            Get
                Return oOCEInfo.Escalation
            End Get
            Set(ByVal Value As Integer)
                oOCEInfo.Escalation = Value
            End Set
        End Property
        Public Property PolicyAmount() As Decimal
            Get
                Return oOCEInfo.PolicyAmount
            End Get
            Set(ByVal Value As Decimal)
                oOCEInfo.PolicyAmount = Value
            End Set
        End Property
        Public Property OverRideAmount() As Decimal
            Get
                Return oOCEInfo.OverRideAmount
            End Get
            Set(ByVal Value As Decimal)
                oOCEInfo.OverRideAmount = Value
            End Set
        End Property
        Public Property SettlementAmount() As Decimal
            Get
                Return oOCEInfo.SettlementAmount
            End Get
            Set(ByVal Value As Decimal)
                oOCEInfo.SettlementAmount = Value
            End Set
        End Property
        Public Property PaidAmount() As Decimal
            Get
                Return oOCEInfo.PaidAmount
            End Get
            Set(ByVal Value As Decimal)
                oOCEInfo.PaidAmount = Value
            End Set
        End Property
        Public Property DateReceived() As Date
            Get
                Return oOCEInfo.DateReceived
            End Get
            Set(ByVal Value As Date)
                oOCEInfo.DateReceived = Value
            End Set
        End Property
        Public Property WorkShopDate() As Date
            Get
                Return oOCEInfo.WorkShopDate
            End Get
            Set(ByVal Value As Date)
                oOCEInfo.WorkShopDate = Value
            End Set
        End Property
        Public Property WorkShopResult() As Integer
            Get
                Return oOCEInfo.WorkShopResult
            End Get
            Set(ByVal Value As Integer)
                oOCEInfo.WorkShopResult = Value
            End Set
        End Property

        Public Property AdminHearingDate() As Date
            Get
                Return oOCEInfo.AdminHearingDate
            End Get
            Set(ByVal Value As Date)
                oOCEInfo.AdminHearingDate = Value
            End Set
        End Property

        Public Property AdminHearingResult() As Integer
            Get
                Return oOCEInfo.AdminHearingResult
            End Get
            Set(ByVal Value As Integer)
                oOCEInfo.AdminHearingResult = Value
            End Set
        End Property
        Public Property WorkshopRequired() As Boolean
            Get
                Return oOCEInfo.WorkshopRequired
            End Get
            Set(ByVal Value As Boolean)
                oOCEInfo.WorkshopRequired = Value
            End Set
        End Property
        Public Property ShowCauseDate() As Date
            Get
                Return oOCEInfo.ShowCauseDate
            End Get
            Set(ByVal Value As Date)
                oOCEInfo.ShowCauseDate = Value
            End Set
        End Property
        Public Property ShowCauseResult() As Integer
            Get
                Return oOCEInfo.ShowCauseResult
            End Get
            Set(ByVal Value As Integer)
                oOCEInfo.ShowCauseResult = Value
            End Set
        End Property
        Public Property CommissionDate() As Date
            Get
                Return oOCEInfo.CommissionDate
            End Get
            Set(ByVal Value As Date)
                oOCEInfo.CommissionDate = Value
            End Set
        End Property
        Public Property CommissionResult() As Integer
            Get
                Return oOCEInfo.CommissionResult
            End Get
            Set(ByVal Value As Integer)
                oOCEInfo.CommissionResult = Value
            End Set
        End Property
        Public Property AgreedOrder() As String
            Get
                Return oOCEInfo.AgreedOrder
            End Get
            Set(ByVal Value As String)
                oOCEInfo.AgreedOrder = Value
            End Set
        End Property
        Public Property AdministrativeOrder() As String
            Get
                Return oOCEInfo.AdministrativeOrder
            End Get
            Set(ByVal Value As String)
                oOCEInfo.AdministrativeOrder = Value
            End Set
        End Property
        Public Property PendingLetter() As Integer
            Get
                Return oOCEInfo.PendingLetter
            End Get
            Set(ByVal Value As Integer)
                oOCEInfo.PendingLetter = Value
            End Set
        End Property
        Public Property LetterGenerated() As Date
            Get
                Return oOCEInfo.LetterGenerated
            End Get
            Set(ByVal Value As Date)
                oOCEInfo.LetterGenerated = Value
            End Set
        End Property
        Public Property RedTagDate() As Date
            Get
                Return oOCEInfo.RedTagDate
            End Get
            Set(ByVal Value As Date)
                oOCEInfo.RedTagDate = Value
            End Set
        End Property
        Public Property LetterPrinted() As Boolean
            Get
                Return oOCEInfo.LetterPrinted
            End Get
            Set(ByVal Value As Boolean)
                oOCEInfo.LetterPrinted = Value
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oOCEInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oOCEInfo.Deleted = Value
            End Set
        End Property
        Public Property CreatedBy() As String
            Get
                Return oOCEInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oOCEInfo.CreatedBy = Value
            End Set
        End Property

        Public Property Comments() As String
            Get
                Return oOCEInfo.Comments
            End Get
            Set(ByVal Value As String)
                oOCEInfo.Comments = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oOCEInfo.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oOCEInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oOCEInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oOCEInfo.ModifiedOn
            End Get
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oOCEInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oOCEInfo.IsDirty = value
            End Set
        End Property
        Public ReadOnly Property colIsDirty() As Boolean
            Get
                Dim xOwnerComplianceEventinfo As MUSTER.Info.OwnerComplianceEventInfo
                For Each xOwnerComplianceEventinfo In colOCE.Values
                    If xOwnerComplianceEventinfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
        End Property
        Public Property EscalationString() As String
            Get
                Return oOCEInfo.EscalationString
            End Get
            Set(ByVal Value As String)
                oOCEInfo.EscalationString = Value
            End Set
        End Property
        Public Property PendingLetterTemplateNum() As Integer
            Get
                Return oOCEInfo.PendingLetterTemplateNum
            End Get
            Set(ByVal Value As Integer)
                oOCEInfo.PendingLetterTemplateNum = Value
            End Set
        End Property
        Public Property EnsiteID() As Integer
            Get
                Return oOCEInfo.EnsiteID
            End Get
            Set(ByVal Value As Integer)
                oOCEInfo.EnsiteID = Value
            End Set
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        Public Sub Load(ByRef ds As DataSet)
            Dim dr As DataRow
            Try
                If ds.Tables("OCE").Rows.Count > 0 Then
                    For Each dr In ds.Tables("OCE").Rows
                        oOCEInfo = New MUSTER.Info.OwnerComplianceEventInfo(dr)
                        colOCE.Add(OCEInfo)
                    Next
                End If
                ds.Tables.Remove("OCE")
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function Retrieve(Optional ByVal id As Integer = 0, Optional ByVal ownerID As Int64 = 0, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.OwnerComplianceEventInfo
            Try
                Dim oOCEInfoLocal As MUSTER.Info.OwnerComplianceEventInfo

                If id = 0 And ownerID = 0 Then
                    Add(id, ownerID, showDeleted)
                ElseIf id <> 0 Then
                    ' check in collection
                    oOCEInfo = colOCE.Item(id)
                ElseIf ownerID <> 0 Then
                    For Each oOCEInfoLocal In colOCE.Values
                        If oOCEInfoLocal.OwnerID = ownerID Then
                            oOCEInfo = oOCEInfoLocal
                            Exit For
                        End If
                    Next
                    Add(id, ownerID, showDeleted)
                End If
                If oOCEInfo Is Nothing Then
                    Add(id, ownerID, showDeleted)
                End If
                Return oOCEInfo
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function Save(ByVal flagNFA As Short, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal bolValidated As Boolean = False, Optional ByVal bolDelete As Boolean = False) As Boolean
            Dim oldID As Integer
            Try
                If Not (oOCEInfo.ID < 0 And oOCEInfo.Deleted) Then
                    oldID = oOCEInfo.ID
                    oOCEDB.Put(flagNFA, oOCEInfo, moduleID, staffID, returnVal)
                    If Not bolValidated Then
                        If oldID < 0 Then
                            colOCE.ChangeKey(oldID, oOCEInfo.ID)
                        End If
                    End If
                    oOCEInfo.Archive()
                    oOCEInfo.IsDirty = False
                End If
                RaiseEvent OwnerComplianceEventChanged(oOCEInfo.IsDirty)
                Return True
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetCAEOverrideAmountHistory(ByVal oceID As Integer) As String
            Dim dsHistory As New DataSet
            Dim strReturn As String = ""
            Dim strSQL As String
            Dim rwData As DataRow

            Try
                strSQL = "exec spGetCAEOverrideAmountHistory " + oceID.ToString
                dsHistory = oOCEDB.DBGetDS(strSQL)
                If dsHistory.Tables(0).Rows.Count > 0 Then
                    For Each rwData In dsHistory.Tables(0).Rows
                        strReturn += IIf(rwData("OverrideAmount") Is DBNull.Value, "NULL", rwData("OverrideAmount")) & "     " & rwData("BeginDate") & "  -  " & rwData("EndDate") & vbCrLf
                    Next
                End If
                Return strReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
#End Region
#Region "Collection Operations"
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal id As Int64, Optional ByVal ownID As Int64 = 0, Optional ByVal showDeleted As Boolean = False)
            Try
                Dim colOCELocal As MUSTER.Info.OwnerComplianceEventsCollection = oOCEDB.DBGetByID(id, ownID, showDeleted)
                If colOCELocal.Count = 0 Then
                    oOCEInfo = New MUSTER.Info.OwnerComplianceEventInfo
                    oOCEInfo.ID = nID
                    oOCEInfo.OwnerID = ownID
                    nID -= 1
                    colOCE.Add(oOCEInfo)
                Else
                    For Each oOCEInfoLocal As MUSTER.Info.OwnerComplianceEventInfo In colOCELocal.Values
                        oOCEInfo = oOCEInfoLocal
                        If oOCEInfo.ID = 0 Then
                            oOCEInfo.ID = nID
                            oOCEInfo.OwnerID = ownID
                            nID -= 1
                        End If
                        colOCE.Add(oOCEInfo)
                    Next
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oOwnerComplianceEvent As MUSTER.Info.OwnerComplianceEventInfo)
            Try
                oOCEInfo = oOwnerComplianceEvent
                If oOCEInfo.ID = 0 Then
                    oOCEInfo.ID = nID
                    nID -= 1
                End If
                colOCE.Add(oOCEInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Integer)
            Try
                If colOCE.Contains(ID) Then
                    colOCE.Remove(ID)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oOwnerComplianceEvent As MUSTER.Info.OwnerComplianceEventInfo)
            Try
                If colOCE.Contains(oOwnerComplianceEvent) Then
                    colOCE.Remove(oOwnerComplianceEvent)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try
                Dim IDs As New Collection
                Dim index As Integer
                Dim xOwnerComplianceEventInfo As MUSTER.Info.OwnerComplianceEventInfo
                For Each xOwnerComplianceEventInfo In colOCE.Values
                    If xOwnerComplianceEventInfo.IsDirty Then
                        oOCEInfo = xOwnerComplianceEventInfo
                        If oOCEInfo.ID < 0 And _
                            Not oOCEInfo.Deleted Then
                            IDs.Add(oOCEInfo.ID)
                        End If
                        Me.Save(moduleID, staffID, returnVal, True)
                    End If
                Next
                If Not (IDs Is Nothing) Then
                    For index = 1 To IDs.Count
                        Dim colKey As String = CType(IDs.Item(index), String)
                        xOwnerComplianceEventInfo = colOCE.Item(colKey)
                        colOCE.ChangeKey(colKey, xOwnerComplianceEventInfo.ID)
                    Next
                End If
                RaiseEvent OwnerComplianceEventChanged(oOCEInfo.IsDirty)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "General Operations"
        Public Sub Clear()
            oOCEInfo = New MUSTER.Info.OwnerComplianceEventInfo
        End Sub
        Public Sub Reset()
            oOCEInfo.Reset()
        End Sub
#End Region
#Region "LookUp Operations"
        Public Function GetWorkshopResults() As DataSet
            Try
                Return oOCEDB.DBGetDS("SELECT PROPERTY_ID, PROPERTY_NAME FROM vCAEProperty WHERE PROPERTY_TYPE_ID = 132 ORDER BY PROPERTY_NAME")
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetShowCauseHearingResults() As DataSet
            Try
                Return oOCEDB.DBGetDS("SELECT PROPERTY_ID, PROPERTY_NAME FROM vCAEProperty WHERE PROPERTY_TYPE_ID = 129 ORDER BY PROPERTY_NAME")
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetCommissionHearingResults() As DataSet
            Try
                Return oOCEDB.DBGetDS("SELECT PROPERTY_ID, PROPERTY_NAME FROM vCAEProperty WHERE PROPERTY_TYPE_ID = 130 ORDER BY PROPERTY_NAME")
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function GetAdminHearingResults() As DataSet
            Try
                Return oOCEDB.DBGetDS("SELECT PROPERTY_ID, PROPERTY_NAME FROM vCAEProperty WHERE PROPERTY_TYPE_ID = 167 ORDER BY PROPERTY_NAME")
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function OwnersPriorViolations(ByVal ownerID As Integer, Optional ByVal excludeOceID As Integer = 0) As DataSet
            Try
                Return oOCEDB.DBOwnersPriorViolations(ownerID, excludeOceID)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetPenaltyByOwnerSizeCitCategory(ByVal strOwnerSize As String, ByVal strCitCategory As String) As Integer
            Dim strSQL As String = String.Empty
            Dim ds As DataSet
            Dim penalty As Integer = 0
            Try
                strSQL = "select top 1 small, medium, large from tblCAE_CITATION_PENALTY where category = '" + strCitCategory + "'"
                ds = oOCEDB.DBGetDS(strSQL)
                If ds.Tables(0).Rows.Count > 0 Then
                    If ds.Tables(0).Columns.Contains(strOwnerSize) Then
                        penalty = ds.Tables(0).Rows(0)(strOwnerSize)
                    End If
                End If
                Return penalty
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetOwnerSize(ByVal ownerID As Integer) As String
            Try
                Return oOCEDB.DBGetOwnerSize(ownerID)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetCitationCategoryPolicyPenalties() As DataSet
            Try
                Return oOCEDB.DBGetDS("SELECT DISTINCT CATEGORY, SMALL, MEDIUM, LARGE FROM tblCAE_CITATION_PENALTY")
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetEnforcements(Optional ByVal ownerID As Integer = 0, Optional ByVal showDeleted As Boolean = False, Optional ByVal returnOCEs As Boolean = False, Optional ByVal oceStatus As Integer = 0, Optional ByVal facility_id As Integer = 0, Optional ByVal UseAllOwnerFormat As Boolean = False, Optional ByVal managerID As Integer = Nothing) As DataSet
            Dim dsInspectionDetails As New DataSet
            Dim dsRel1 As DataRelation
            Dim dsRel2 As DataRelation
            Dim dsRel3 As DataRelation
            Try
                dsInspectionDetails = oOCEDB.DBGetEnforcements(ownerID, showDeleted, returnOCEs, oceStatus, facility_id, UseAllOwnerFormat, managerID)

                If dsInspectionDetails.Tables.Count <= 0 Then
                    Return dsInspectionDetails
                End If

                Dim parentCols(3) As DataColumn
                Dim childCols(3) As DataColumn

                If ownerID > 0 And Not UseAllOwnerFormat Then
                    dsRel2 = New DataRelation("OCEToFCE", dsInspectionDetails.Tables(0).Columns("OCE_ID"), dsInspectionDetails.Tables(1).Columns("OCE_ID"), False)
                    dsRel3 = New DataRelation("InspectionCitationToDiscrep", dsInspectionDetails.Tables(1).Columns("INS_CIT_ID"), dsInspectionDetails.Tables(2).Columns("INS_CIT_ID"), False)

                    dsInspectionDetails.Relations.Add(dsRel2)
                    dsInspectionDetails.Relations.Add(dsRel3)
                Else
                    dsRel1 = New DataRelation("StatusToOwner", dsInspectionDetails.Tables(0).Columns("OCE_STATUS"), dsInspectionDetails.Tables(1).Columns("OCE_STATUS"), False)
                    dsRel2 = New DataRelation("OCEToFCE", dsInspectionDetails.Tables(1).Columns("OCE_ID"), dsInspectionDetails.Tables(2).Columns("OCE_ID"), False)
                    dsRel3 = New DataRelation("InspectionCitationToDiscrep", dsInspectionDetails.Tables(2).Columns("INS_CIT_ID"), dsInspectionDetails.Tables(3).Columns("INS_CIT_ID"), False)

                    dsInspectionDetails.Relations.Add(dsRel1)
                    dsInspectionDetails.Relations.Add(dsRel2)
                    dsInspectionDetails.Relations.Add(dsRel3)
                End If

                Return dsInspectionDetails
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetPriorEnforcements(Optional ByVal ownerID As Integer = 0, Optional ByVal showDeleted As Boolean = False) As DataSet
            Dim dsInspectionDetails As New DataSet
            Dim dsRel1 As DataRelation
            Dim dsRel2 As DataRelation
            Dim dsRel3 As DataRelation
            Try
                dsInspectionDetails = oOCEDB.DBGetPriorEnforcements(ownerID, showDeleted)
                If dsInspectionDetails.Tables.Count <= 0 Then
                    Return dsInspectionDetails
                End If

                dsRel1 = New DataRelation("StatusToOwner", dsInspectionDetails.Tables(0).Columns("OCE_STATUS"), dsInspectionDetails.Tables(1).Columns("OCE_STATUS"), False)
                dsRel2 = New DataRelation("OwnerToCitation", dsInspectionDetails.Tables(1).Columns("OCE_ID"), dsInspectionDetails.Tables(2).Columns("OCE_ID"), False)
                dsRel3 = New DataRelation("CitationToDiscrep", dsInspectionDetails.Tables(2).Columns("INS_CIT_ID"), dsInspectionDetails.Tables(3).Columns("INS_CIT_ID"), False)
                dsInspectionDetails.Relations.Add(dsRel1)
                dsInspectionDetails.Relations.Add(dsRel2)
                dsInspectionDetails.Relations.Add(dsRel3)

                Return dsInspectionDetails
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function GetMyOCERedTagChanges(ByVal OCEChangedDate As Date, ByVal user As String) As DataSet
            Dim dsOCE As New DataSet
            Try

                dsOCE = oOCEDB.DBGetMyRedTagStatusChanges(OCEChangedDate, user)

                Return dsOCE

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function GetCOIFacs(ByVal oceID As Integer, Optional ByVal showDeleted As Boolean = False) As DataSet
            Try
                Return oOCEDB.DBGetCOIFacs(oceID, showDeleted)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetCAEOCEEscalation(ByVal flagNFA As Short, _
                                        ByVal oceID As Integer, _
                                        ByVal oceStatus As Integer, _
                                        ByVal ownerID As Integer, _
                                        ByVal nextDueDate As Date, _
                                        ByVal overrideDueDate As Date, _
                                        ByVal policyAmount As Decimal, _
                                        ByVal overrideAmount As Decimal, _
                                        ByVal settlementAmount As Decimal, _
                                        ByVal paidAmount As Decimal, _
                                        ByVal workshopRequired As Boolean, _
                                        ByVal workshopDate As Date, _
                                        ByVal workshopResult As Integer, _
                                        ByVal showCauseDate As Date, _
                                        ByVal showCauseResult As Integer, _
                                        ByVal commissionDate As Date, _
                                        ByVal commissionResult As Integer, _
                                        ByVal pendingLetter As Integer, _
                                        ByVal citationDueDate As Date, _
                                        ByVal ocePath As Integer, _
                                        ByVal dateReceived As Date, _
                                        ByRef escalationID As Integer, ByVal adminDate As Object, ByVal adminResult As Object) As String
            Try
                Return oOCEDB.DBGetCAEOCEEscalation(flagNFA, _
                                                    oceID, _
                                                    oceStatus, _
                                                    ownerID, _
                                                    nextDueDate, _
                                                    overrideDueDate, _
                                                    policyAmount, _
                                                    overrideAmount, _
                                                    settlementAmount, _
                                                    paidAmount, _
                                                    workshopRequired, _
                                                    workshopDate, _
                                                    workshopResult, _
                                                    showCauseDate, _
                                                    showCauseResult, _
                                                    commissionDate, _
                                                    commissionResult, _
                                                    pendingLetter, _
                                                    citationDueDate, _
                                                    ocePath, _
                                                    dateReceived, _
                                                    escalationID, adminDate, adminResult)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function ExecuteCAEOCEEscalation(ByVal flagNFA As Short, _
                                                ByVal oceID As Integer, _
                                                ByRef oceStatus As Integer, _
                                                ByVal ownerID As Integer, _
                                                ByRef nextDueDate As Date, _
                                                ByRef overrideDueDate As Date, _
                                                ByVal policyAmount As Decimal, _
                                                ByVal overrideAmount As Decimal, _
                                                ByRef settlementAmount As Decimal, _
                                                ByVal paidAmount As Decimal, _
                                                ByVal workshopRequired As Boolean, _
                                                ByVal workshopDate As Date, _
                                                ByVal workshopResult As Integer, _
                                                ByRef showCauseDate As Date, _
                                                ByRef showCauseResult As Integer, _
                                                ByRef commissionDate As Date, _
                                                ByRef commissionResult As Integer, _
                                                ByRef pendingLetter As Integer, _
                                                ByVal citationDueDate As Date, _
                                                ByVal ocePath As Integer, _
                                                ByVal dateReceived As Date, _
                                                ByVal userProvidedDate As Date, _
                                                ByRef stroceStatus As String, _
                                                ByRef strpendingLetter As String, _
                                                ByRef strEscalation As String, _
                                                ByRef pendingLetterTemplateNum As Integer, _
                                                ByRef escalationID As Integer, ByVal adminDate As Object, ByVal adminResult As Object, ByVal userID As String) As String
            Try
                Return oOCEDB.DBExecuteCAEOCEEscalation(flagNFA, _
                                                    oceID, _
                                                    oceStatus, _
                                                    ownerID, _
                                                    nextDueDate, _
                                                    overrideDueDate, _
                                                    policyAmount, _
                                                    overrideAmount, _
                                                    settlementAmount, _
                                                    paidAmount, _
                                                    workshopRequired, _
                                                    workshopDate, _
                                                    workshopResult, _
                                                    showCauseDate, _
                                                    showCauseResult, _
                                                    commissionDate, _
                                                    commissionResult, _
                                                    pendingLetter, _
                                                    citationDueDate, _
                                                    ocePath, _
                                                    dateReceived, _
                                                    userProvidedDate, _
                                                    stroceStatus, _
                                                    strpendingLetter, _
                                                    strEscalation, _
                                                    pendingLetterTemplateNum, _
                                                    escalationID, adminDate, adminResult, userID)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetLetterGeneratedDate(ByVal entityID As Integer, ByVal entityType As Integer, Optional ByVal propertyIDTemplateNum As Integer = 0, Optional ByVal showDeleted As Boolean = False) As Date
            Try
                Return oOCEDB.DBGetLetterGeneratedDate(entityID, entityType, propertyIDTemplateNum, showDeleted)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Sub SaveLetterGeneratedDate(ByRef letterGenID As Integer, ByVal entityID As Integer, ByVal entityType As Integer, ByVal propertyIDTemplateNum As Integer, ByVal generatedDate As Date, ByVal deleted As Boolean, ByVal documentID As Integer, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try
                oOCEDB.DBPutLetterGeneratedDate(letterGenID, entityID, entityType, propertyIDTemplateNum, generatedDate, deleted, documentID, moduleID, staffID, returnVal)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function OwnerHasPrevWorkshopDate(ByVal ownerID As Integer, Optional ByVal showDeleted As Boolean = False, Optional ByVal excludeOceID As Integer = 0) As Boolean
            Try
                Return oOCEDB.DBGetCAEPrevWorkshopDate(ownerID, showDeleted, excludeOceID)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetOpenOCEDocumentInfoforOwner(ByVal ownerID As Integer) As String
            Try
                Return oOCEDB.DBGetOpenOCEDocumentInfoforOwner(ownerID)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function RunSQLQuery(ByVal strSQL As String) As DataSet
            Try
                Return oOCEDB.DBGetDS(strSQL)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function OwnerHasWorkshopOCEDuringPast90Days(ByVal ownerID As Integer, Optional ByVal excludeOceID As Integer = 0) As Boolean
            Try
                Return oOCEDB.DBOwnerHasWorkshopOCEDuringPast90Days(ownerID, excludeOceID)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region
#End Region
#Region "External Event Handlers"
        Private Sub OwnerComplianceEventInfoChanged(ByVal bolValue As Boolean) Handles oOCEInfo.OwnerComplianceEventInfoChanged
            RaiseEvent OwnerComplianceEventChanged(bolValue)
        End Sub
        Private Sub OwnerComplianceEventColChanged(ByVal bolValue As Boolean) Handles colOCE.OwnerComplianceEventColChanged
            RaiseEvent ColChanged(bolValue)
        End Sub
#End Region
    End Class
End Namespace
