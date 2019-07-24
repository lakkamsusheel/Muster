'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.FacilityComplianceEvent
'   Provides the operations required to manipulate an Entity object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0                              Original class definition
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
' NOTE: This file to be used as FacilityComplianceEvent to build other objects.
'       Replace keyword "FacilityComplianceEvent" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pFacilityComplianceEvent
#Region "Public Events"
        Public Event FacilityComplianceEventErr(ByVal MsgStr As String)
        Public Event FacilityComplianceEventChanged(ByVal bolValue As Boolean)
        Public Event ColChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        Private WithEvents oFCEInfo As MUSTER.Info.FacilityComplianceEventInfo
        Private WithEvents colFCE As MUSTER.Info.FacilityComplianceEventsCollection
        Private oFCEDB As MUSTER.DataAccess.FacilityComplianceEventDB
        Private MusterException As New MUSTER.Exceptions.MusterExceptions

        'Private CAEFceCitations As New MUSTER.BusinessLogic.pCAEFceCitations
        'Private CitationPenalty As New MUSTER.BusinessLogic.pCitationPenalty
        'Private pInsCitation As New MUSTER.BusinessLogic.pInspectionCitation

        Private nID As Int64 = -1
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
#End Region
#Region "Constructors"
        Public Sub New(Optional ByVal strDBConn As String = "", Optional ByRef MusterXCEP As MUSTER.Exceptions.MusterExceptions = Nothing)
            If MusterXCEP Is Nothing Then
                MusterException = New MUSTER.Exceptions.MusterExceptions
            Else
                MusterException = MusterXCEP
            End If
            oFCEInfo = New MUSTER.Info.FacilityComplianceEventInfo
            colFCE = New MUSTER.Info.FacilityComplianceEventsCollection
            oFCEDB = New MUSTER.DataAccess.FacilityComplianceEventDB(strDBConn, MusterXCEP)
        End Sub
#End Region
#Region "Exposed Attributes"
        'Public Property FCECitations() As MUSTER.BusinessLogic.pCitationPenalty
        '    Get
        '        Return CitationPenalty
        '    End Get
        '    Set(ByVal Value As MUSTER.BusinessLogic.pCitationPenalty)
        '        CitationPenalty = Value
        '    End Set
        'End Property
        'Public Property InsCitation() As MUSTER.BusinessLogic.pInspectionCitation
        '    Get
        '        Return pInsCitation
        '    End Get
        '    Set(ByVal Value As MUSTER.BusinessLogic.pInspectionCitation)
        '        pInsCitation = Value
        '    End Set
        'End Property
        'Public Property OwnerName() As String
        '    Get
        '        Return oFCEInfo.OwnerName
        '    End Get
        '    Set(ByVal Value As String)
        '        oFCEInfo.OwnerName = Value
        '    End Set
        'End Property
        'Public Property FacilityName() As String
        '    Get
        '        Return oFCEInfo.FacilityName
        '    End Get
        '    Set(ByVal Value As String)
        '        oFCEInfo.FacilityName = Value
        '    End Set
        'End Property
        'Public Property InspectorName() As String
        '    Get
        '        Return oFCEInfo.InspectorName
        '    End Get
        '    Set(ByVal Value As String)
        '        oFCEInfo.InspectorName = Value
        '    End Set
        'End Property
        'Public Property InspectedOn() As DateTime
        '    Get
        '        Return oFCEInfo.InspectedOn
        '    End Get
        '    Set(ByVal Value As DateTime)
        '        oFCEInfo.InspectedOn = Value
        '    End Set
        'End Property
        'Public Property Citations() As Integer
        '    Get
        '        Return oFCEInfo.Citations
        '    End Get
        '    Set(ByVal Value As Integer)
        '        oFCEInfo.Citations = Value
        '    End Set
        'End Property

        Public ReadOnly Property ID() As Int32
            Get
                Return oFCEInfo.ID
            End Get
        End Property
        Public Property InspectionID() As Int32
            Get
                Return oFCEInfo.InspectionID
            End Get
            Set(ByVal Value As Int32)
                oFCEInfo.InspectionID = Value
            End Set
        End Property
        Public Property OwnerID() As Int32
            Get
                Return oFCEInfo.OwnerID
            End Get
            Set(ByVal Value As Int32)
                oFCEInfo.OwnerID = Value
            End Set
        End Property
        Public Property FacilityID() As Int32
            Get
                Return oFCEInfo.FacilityID
            End Get
            Set(ByVal Value As Int32)
                oFCEInfo.FacilityID = Value
            End Set
        End Property
        Public Property FCEDate() As DateTime
            Get
                Return oFCEInfo.FCEDate
            End Get
            Set(ByVal Value As DateTime)
                oFCEInfo.FCEDate = Value
            End Set
        End Property
        Public Property Source() As String
            Get
                Return oFCEInfo.Source
            End Get
            Set(ByVal Value As String)
                oFCEInfo.Source = Value
            End Set
        End Property
        Public Property DueDate() As DateTime
            Get
                Return oFCEInfo.DueDate
            End Get
            Set(ByVal Value As DateTime)
                oFCEInfo.DueDate = Value
            End Set
        End Property
        Public Property ReceivedDate() As DateTime
            Get
                Return oFCEInfo.ReceivedDate
            End Get
            Set(ByVal Value As DateTime)
                oFCEInfo.ReceivedDate = Value
            End Set
        End Property
        Public Property CreatedBy() As String
            Get
                Return oFCEInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oFCEInfo.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oFCEInfo.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oFCEInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oFCEInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oFCEInfo.ModifiedOn
            End Get
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oFCEInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oFCEInfo.Deleted = Value
            End Set
        End Property
        Public Property OCEGenerated() As Boolean
            Get
                Return oFCEInfo.OCEGenerated
            End Get
            Set(ByVal Value As Boolean)
                oFCEInfo.OCEGenerated = Value
            End Set
        End Property
        Public Property OCEID() As Int32
            Get
                Return oFCEInfo.OCEID
            End Get
            Set(ByVal Value As Int32)
                oFCEInfo.OCEID = Value
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oFCEInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oFCEInfo.IsDirty = value
            End Set
        End Property
        Public ReadOnly Property colIsDirty() As Boolean
            Get
                Dim xFCEinfo As MUSTER.Info.FacilityComplianceEventInfo
                For Each xFCEinfo In colFCE.Values
                    If xFCEinfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        Public Function Retrieve(Optional ByVal id As Integer = 0, Optional ByVal inspID As Int64 = 0, Optional ByVal ownerID As Int64 = 0, Optional ByVal facID As Int64 = 0, Optional ByVal showDeleted As Boolean = False, Optional ByVal bolLoading As Boolean = False) As MUSTER.Info.FacilityComplianceEventInfo
            Dim bolDataAged As Boolean = False
            Try
                Dim oFCEInfoLocal As MUSTER.Info.FacilityComplianceEventInfo

                If id = 0 And inspID = 0 And ownerID = 0 And facID = 0 Then
                    Add(id, inspID, ownerID, facID, showDeleted)
                ElseIf id <> 0 Then
                    ' check in collection
                    oFCEInfo = colFCE.Item(id)
                ElseIf inspID <> 0 Then
                    For Each oFCEInfoLocal In colFCE.Values
                        If oFCEInfoLocal.InspectionID = inspID Then
                            oFCEInfo = oFCEInfoLocal
                            Exit For
                        End If
                    Next
                    Add(id, inspID, ownerID, facID, showDeleted)
                ElseIf facID <> 0 Then
                    For Each oFCEInfoLocal In colFCE.Values
                        If oFCEInfoLocal.FacilityID = facID Then
                            oFCEInfo = oFCEInfoLocal
                            Exit For
                        End If
                    Next
                    Add(id, inspID, ownerID, facID, showDeleted)
                ElseIf ownerID <> 0 Then
                    For Each oFCEInfoLocal In colFCE.Values
                        If oFCEInfoLocal.OwnerID = ownerID Then
                            oFCEInfo = oFCEInfoLocal
                            Exit For
                        End If
                    Next
                    Add(id, inspID, ownerID, facID, showDeleted)
                End If

                ' Check for Aged Data here.
                'If Not (oFCEInfo Is Nothing) Then
                '    If oFCEInfo.IsAgedData = True And oFCEInfo.IsDirty = False Then
                '        bolDataAged = True
                '        colFCE.Remove(oFCEInfo)
                '    End If
                'End If

                If oFCEInfo Is Nothing Then 'Or bolDataAged Then
                    Add(id, inspID, ownerID, facID, showDeleted)
                End If
                Return oFCEInfo
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal bolValidated As Boolean = False, Optional ByVal bolDelete As Boolean = False, Optional ByVal OverrideRights As Boolean = False) As Boolean
            Dim oldID As Integer
            Try
                If Not (oFCEInfo.ID < 0 And oFCEInfo.Deleted) Then
                    oldID = oFCEInfo.ID
                    oFCEDB.Put(oFCEInfo, moduleID, staffID, returnVal, OverrideRights)
                    If Not returnVal = String.Empty Then
                        Exit Function
                    End If

                    If Not bolValidated Then
                        If oldID < 0 Then
                            colFCE.ChangeKey(oldID, oFCEInfo.ID)
                        End If
                    End If
                    oFCEInfo.Archive()
                    oFCEInfo.IsDirty = False
                End If
                If Not bolValidated And bolDelete Then
                    If oFCEInfo.Deleted Then
                        ' check if other fce's are present else load new instance
                        Dim strNext As String = Me.GetNext()
                        Dim strPrev As String = Me.GetPrevious()
                        If strNext = oFCEInfo.ID Then
                            If strPrev = oFCEInfo.ID Then
                                RaiseEvent FacilityComplianceEventErr("FCE " + oFCEInfo.ID.ToString + " deleted")
                                colFCE.Remove(oFCEInfo)
                                If bolDelete Then
                                    oFCEInfo = New MUSTER.Info.FacilityComplianceEventInfo
                                Else
                                    oFCEInfo = Me.Retrieve(0)
                                End If
                            Else
                                RaiseEvent FacilityComplianceEventErr("FCE " + oFCEInfo.ID.ToString + " deleted")
                                colFCE.Remove(oFCEInfo)
                                oFCEInfo = Me.Retrieve(strPrev)
                            End If
                        Else
                            RaiseEvent FacilityComplianceEventErr("FCE " + oFCEInfo.ID.ToString + " deleted")
                            colFCE.Remove(oFCEInfo)
                            oFCEInfo = Me.Retrieve(strNext)
                        End If
                    End If
                End If
                RaiseEvent FacilityComplianceEventChanged(oFCEInfo.IsDirty)
                Return True
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
#End Region
#Region "Collection Operations"
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal id As Int64, Optional ByVal inspID As Int64 = 0, Optional ByVal ownID As Int64 = 0, Optional ByVal facID As Int64 = 0, Optional ByVal showDeleted As Boolean = False)
            Try
                Dim colFCELocal As MUSTER.Info.FacilityComplianceEventsCollection = oFCEDB.DBGetByID(id, inspID, ownID, facID, showDeleted)
                If colFCELocal.Count = 0 Then
                    oFCEInfo = New MUSTER.Info.FacilityComplianceEventInfo
                    oFCEInfo.ID = nID
                    oFCEInfo.InspectionID = inspID
                    oFCEInfo.OwnerID = ownID
                    oFCEInfo.FacilityID = facID
                    nID -= 1
                    colFCE.Add(oFCEInfo)
                Else
                    For Each oFCEInfoLocal As MUSTER.Info.FacilityComplianceEventInfo In colFCELocal.Values
                        oFCEInfo = oFCEInfoLocal
                        If oFCEInfo.ID = 0 Then
                            oFCEInfo.ID = nID
                            oFCEInfo.InspectionID = inspID
                            oFCEInfo.OwnerID = ownID
                            oFCEInfo.FacilityID = facID
                            nID -= 1
                        End If
                        colFCE.Add(oFCEInfo)
                    Next
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oFacilityComplianceEvent As MUSTER.Info.FacilityComplianceEventInfo)
            Try
                oFCEInfo = oFacilityComplianceEvent
                If oFCEInfo.ID = 0 Then
                    oFCEInfo.ID = nID
                    nID -= 1
                End If
                colFCE.Add(oFCEInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Integer)
            Try
                If colFCE.Contains(ID) Then
                    colFCE.Remove(ID)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oFacilityComplianceEvent As MUSTER.Info.FacilityComplianceEventInfo)
            Try
                If colFCE.Contains(oFacilityComplianceEvent) Then
                    colFCE.Remove(oFacilityComplianceEvent)
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
                Dim xFCEInfo As MUSTER.Info.FacilityComplianceEventInfo
                For Each xFCEInfo In colFCE.Values
                    If xFCEInfo.IsDirty Then
                        oFCEInfo = xFCEInfo
                        If oFCEInfo.ID <= 0 And _
                            Not oFCEInfo.Deleted Then
                            IDs.Add(oFCEInfo.ID)
                        End If
                        Me.Save(moduleID, staffID, returnVal, True)
                    End If
                Next
                If Not (IDs Is Nothing) Then
                    For index = 1 To IDs.Count
                        Dim colKey As String = CType(IDs.Item(index), String)
                        xFCEInfo = colFCE.Item(colKey)
                        colFCE.ChangeKey(colKey, xFCEInfo.ID)
                    Next
                End If
                RaiseEvent FacilityComplianceEventChanged(oFCEInfo.IsDirty)
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
            Dim strArr() As String = colFCE.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colFCE.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            ElseIf colFCE.Count <> 0 Then
                Return colFCE.Item(nArr.GetValue(colIndex)).ID.ToString
            Else
                Return "0"
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oFCEInfo = New MUSTER.Info.FacilityComplianceEventInfo
        End Sub
        Public Sub Reset()
            oFCEInfo.Reset()
        End Sub
#End Region
#Region "LookUp Operations"
        'Public Function PopulateFCEDataset() As DataSet
        '    Dim dsFacilityInsEvent As New DataSet
        '    Dim drRow As DataRow
        '    Dim oCol As DataColumn
        '    Dim dsRel As DataRelation
        '    Dim strSQL As String
        '    Try

        '        strSQL = "SELECT * FROM VCAE_FCE_INS_DATA WHERE DELETED = 0 ;" & _
        '                 "SELECT * FROM VCAE_FCE_INS_CITATION_DATA WHERE DELETED = 0 ORDER BY Citation_ID;"

        '        dsFacilityInsEvent = oFCEDB.DBGetDS(strSQL)
        '        If dsFacilityInsEvent.Tables.Count <= 0 Then
        '            Return dsFacilityInsEvent
        '        End If
        '        For Each oCol In dsFacilityInsEvent.Tables(0).Columns
        '            oCol.ReadOnly = True
        '        Next
        '        For Each oCol In dsFacilityInsEvent.Tables(1).Columns
        '            oCol.ReadOnly = True
        '        Next

        '        dsRel = New DataRelation("FCEtoINS", dsFacilityInsEvent.Tables(0).Columns("Inspection_ID"), dsFacilityInsEvent.Tables(1).Columns("Inspection_ID"), False)

        '        dsFacilityInsEvent.Relations.Add(dsRel)
        '        Return dsFacilityInsEvent
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Function
        'Public Function PopulateFacilityName(Optional ByVal OwnerID As Int32 = 0) As DataTable
        '    Try
        '        Dim dtReturn As DataTable = GetDataTable("vCAE_FacilityName", OwnerID)
        '        Return dtReturn
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Function
        'Public Function PopulateOwnerName() As DataTable
        '    Try
        '        Dim dtReturn As DataTable = GetDataTable("v_OWNER_NAME", OwnerID)
        '        Return dtReturn
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Function
        'Public Function PopulateCitationList() As DataSet
        '    Try
        '        Dim dsReturn As DataSet = oFCEDB.GetCitationList(True)
        '        Return dsReturn
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try

        'End Function
        'Private Function GetDataTable(ByVal DBViewName As String, Optional ByVal OwnerID As Int32 = 0) As DataTable
        '    Dim dsReturn As New DataSet
        '    Dim dtReturn As DataTable
        '    Dim strSQL As String
        '    Try
        '        If OwnerID <> 0 Then
        '            strSQL = "SELECT Facility_Name,Facility_ID FROM " & DBViewName & " WHERE OWNER_ID = " & OwnerID.ToString
        '        Else
        '            strSQL = "SELECT * FROM " & DBViewName
        '        End If


        '        dsReturn = oFCEDB.DBGetDS(strSQL)
        '        If dsReturn.Tables(0).Rows.Count > 0 Then
        '            dtReturn = dsReturn.Tables(0)
        '        Else
        '            dtReturn = Nothing
        '        End If
        '        Return dtReturn
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Function
#End Region
#Region "Miscellaneous Operations"
        Public Function GetInspections(Optional ByVal facility_id As Integer = 0, Optional ByVal managerID As Integer = Nothing) As DataSet
            Dim dsInspectionDetails As New DataSet
            Dim dsRel1 As DataRelation
            Dim dsRel2 As DataRelation
            Try
                dsInspectionDetails = oFCEDB.DBGetInspections(False, facility_id, managerID)
                If dsInspectionDetails.Tables.Count <= 0 Then
                    Return dsInspectionDetails
                End If

                dsRel1 = New DataRelation("OwnerToFacility", dsInspectionDetails.Tables(0).Columns("Owner_ID"), dsInspectionDetails.Tables(1).Columns("Owner_ID"), False)
                dsRel2 = New DataRelation("InspectionToCitation", dsInspectionDetails.Tables(1).Columns("INSPECTION_ID"), dsInspectionDetails.Tables(2).Columns("INSPECTION_ID"), False)

                dsInspectionDetails.Relations.Add(dsRel1)
                dsInspectionDetails.Relations.Add(dsRel2)

                Return dsInspectionDetails
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetCompliances(Optional ByVal facID As Integer = 0, Optional ByVal showDeleted As Boolean = False, Optional ByVal managerID As Integer = Nothing) As DataSet
            Dim dsInspectionDetails As New DataSet
            Dim dsRel1 As DataRelation
            Dim dsRel2 As DataRelation
            Dim dsRel3 As DataRelation
            Try
                dsInspectionDetails = oFCEDB.DBGetCompliances(facID, showDeleted, managerID)
                If dsInspectionDetails.Tables.Count <= 0 Then
                    Return dsInspectionDetails
                End If

                Dim parentCols(3) As DataColumn
                Dim childCols(3) As DataColumn

                If facID > 0 Then
                    dsRel2 = New DataRelation("FacilityToInspectionCitation", dsInspectionDetails.Tables(0).Columns("FCE_ID"), dsInspectionDetails.Tables(1).Columns("FCE_ID"), False)
                    parentCols(0) = dsInspectionDetails.Tables(1).Columns("INSPECTION_ID")
                    parentCols(1) = dsInspectionDetails.Tables(1).Columns("CITATION_ID")
                    parentCols(2) = dsInspectionDetails.Tables(1).Columns("QUESTION_ID")
                    parentCols(3) = dsInspectionDetails.Tables(1).Columns("INS_CIT_ID")
                    childCols(0) = dsInspectionDetails.Tables(2).Columns("INSPECTION_ID")
                    childCols(1) = dsInspectionDetails.Tables(2).Columns("CITATION_ID")
                    childCols(2) = dsInspectionDetails.Tables(2).Columns("QUESTION_ID")
                    childCols(3) = dsInspectionDetails.Tables(2).Columns("INS_CIT_ID")
                    dsRel3 = New DataRelation("InspectionCitationToDiscrep", parentCols, childCols, False)
                    dsInspectionDetails.Relations.Add(dsRel2)
                    dsInspectionDetails.Relations.Add(dsRel3)
                Else
                    dsRel1 = New DataRelation("OwnerToFacility", dsInspectionDetails.Tables(0).Columns("Owner_ID"), dsInspectionDetails.Tables(1).Columns("Owner_ID"), False)
                    dsRel2 = New DataRelation("FacilityToInspectionCitation", dsInspectionDetails.Tables(1).Columns("FCE_ID"), dsInspectionDetails.Tables(2).Columns("FCE_ID"), False)
                    parentCols(0) = dsInspectionDetails.Tables(2).Columns("INSPECTION_ID")
                    parentCols(1) = dsInspectionDetails.Tables(2).Columns("CITATION_ID")
                    parentCols(2) = dsInspectionDetails.Tables(2).Columns("QUESTION_ID")
                    parentCols(3) = dsInspectionDetails.Tables(2).Columns("INS_CIT_ID")
                    childCols(0) = dsInspectionDetails.Tables(3).Columns("INSPECTION_ID")
                    childCols(1) = dsInspectionDetails.Tables(3).Columns("CITATION_ID")
                    childCols(2) = dsInspectionDetails.Tables(3).Columns("QUESTION_ID")
                    childCols(3) = dsInspectionDetails.Tables(3).Columns("INS_CIT_ID")
                    dsRel3 = New DataRelation("InspectionCitationToDiscrep", parentCols, childCols, False)
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
        Public Function GetOwners() As DataSet
            Try
                Dim ds As DataSet = oFCEDB.DBGetDS("SELECT * FROM v_OWNER_NAME")
                Return ds
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetFacilities(ByVal ownID As Integer) As DataSet
            Try
                Dim strSQL As String = "SELECT * FROM vCAE_FacilityName WHERE OWNER_ID = " + ownID.ToString + " ORDER BY FACILITY_ID"
                Dim ds As DataSet = oFCEDB.DBGetDS(strSQL)
                Return ds
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetCitations(Optional ByVal strWhere As String = "") As DataSet
            Dim strSQL As String
            Try
                strSQL = "SELECT * FROM tblCAE_CITATION_PENALTY"
                If strWhere <> String.Empty Then
                    strSQL += strWhere
                End If
                Return oFCEDB.DBGetDS(strSQL)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetDiscrepText(Optional ByVal strWhere As String = "") As DataSet
            Dim strSQL As String
            Try
                strSQL = "SELECT CITATION AS CITATION_ID, QUESTION_ID, DISCREP_TEXT AS [DISCREP TEXT] FROM tblINS_INSPECTION_CHECKLIST_MASTER WHERE DELETED = 0 AND DISCREP_TEXT IS NOT NULL AND CITATION = 19"
                If strWhere <> String.Empty Then
                    strSQL += strWhere
                End If
                Return oFCEDB.DBGetDS(strSQL)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        '#Region "OCE creation/modification Logic"
        '        Public Function OCECreationLogic(ByVal OwnerID As Integer, ByVal arrFCEs As ArrayList) As Boolean
        '            Dim i As Integer
        '            Dim InsCitationCol As MUSTER.Info.InspectionCitationsCollection
        '            Dim InsCitationInfo As MUSTER.Info.InspectionCitationInfo
        '            Dim citationInfo As MUSTER.Info.CitationPenaltyInfo
        '            Dim bolDiscrepancy As Boolean = True
        '            Dim bolPriorValidation As Boolean = False
        '            Dim bolCat1 As Boolean = False
        '            Dim bolCat2 As Boolean = False
        '            Try
        '                For i = 0 To arrFCEs.Count - 1
        '                    InsCitationCol = pInsCitation.RetrieveByFCEID(arrFCEs.Item(i))
        '                    For Each InsCitationInfo In InsCitationCol.Values
        '                        ' Does this owner have anu prior violations
        '                        bolPriorValidation = True
        '                        citationInfo = CitationPenalty.Retrieve(InsCitationInfo.CitationID)
        '                        If Not citationInfo.Category.ToUpper = "DISCREPANCY" Then
        '                            bolDiscrepancy = False
        '                        End If
        '                        If citationInfo.Category.Trim = "1" Then
        '                            bolCat1 = True
        '                        End If
        '                        If citationInfo.Category.ToUpper = "2A" Or citationInfo.Category.ToUpper = "2B" Or citationInfo.Category.ToUpper = "2C" Then
        '                            bolCat2 = True
        '                        End If
        '                    Next
        '                Next
        '                If bolDiscrepancy Then
        '                    DiscrepanciesOnly(OwnerID, arrFCEs)
        '                ElseIf bolPriorValidation Then
        '                    PriorViolations()
        '                ElseIf bolCat1 Then
        '                    Cat1NoPriors()
        '                ElseIf bolCat2 Then
        '                    Cat2NoPriors()
        '                Else
        '                    Cat3NoPriors()
        '                End If
        '            Catch ex As Exception
        '                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '                Throw ex
        '            End Try
        '        End Function
        '        Public Sub DiscrepanciesOnly(ByVal OwnerID As Integer, ByVal arrFCEs As ArrayList)
        '            Try
        '                pOwner.Retrieve(OwnerID)
        '                If pOwner.OwnerInfo.facilityCollection.Count < 5 Then
        '                    If TankStatusCIU(pOwner) Then
        '                        ' Execute the Logic for SMALL Owner

        '                    End If
        '                ElseIf pOwner.OwnerInfo.facilityCollection.Count < 15 And pOwner.OwnerInfo.facilityCollection.Count > 5 Then
        '                    If TankStatusCIU(pOwner) Then
        '                        ' Execute the Logic for MEDIUM Owner
        '                    End If
        '                Else
        '                    If TankStatusCIU(pOwner) Then
        '                        'execute the logic for LARGE owner
        '                    End If
        '                End If
        '            Catch ex As Exception
        '                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '                Throw ex
        '            End Try
        '        End Sub
        '        Public Function TankStatusCIU(ByVal pOwner As MUSTER.BusinessLogic.pOwner) As Boolean
        '            Try
        '                For Each FacInfo As MUSTER.info.FacilityInfo In pOwner.OwnerInfo.facilityCollection.Values
        '                    For Each tankInfo As MUSTER.Info.TankInfo In FacInfo.TankCollection.Values
        '                        If tankInfo.TankStatus = 424 Then ' Currently in use - 424
        '                            Return True
        '                        End If
        '                    Next
        '                Next
        '            Catch ex As Exception
        '                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '                Throw ex
        '            End Try
        '        End Function
        '        Public Sub PriorViolations()
        '            Try

        '            Catch ex As Exception
        '                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '                Throw ex
        '            End Try
        '        End Sub
        '        Public Sub Cat1NoPriors()
        '            Try

        '            Catch ex As Exception
        '                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '                Throw ex
        '            End Try
        '        End Sub
        '        Public Sub Cat2NoPriors()
        '            Try

        '            Catch ex As Exception
        '                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '                Throw ex
        '            End Try
        '        End Sub
        '        Public Sub Cat3NoPriors()
        '            Try

        '            Catch ex As Exception
        '                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '                Throw ex
        '            End Try
        '        End Sub
        '#End Region
        'Public Function EntityTable() As DataSet
        '    Dim oFCEInfoLocal As MUSTER.Info.FacilityComplianceEventInfo
        '    Dim oCitationInfoLocal As MUSTER.Info.CitationPenaltyInfo
        '    Dim oFCECitationInfoLocal As MUSTER.Info.CAEFceCitationInfo
        '    Dim dr, ownerRow, CitationRow As DataRow
        '    Dim ds As New DataSet
        '    Dim tbEntityTable As New DataTable
        '    Dim Ownertable As New DataTable
        '    Dim CitationTable As New DataTable
        '    Dim relation1, relation2 As DataRelation
        '    Dim insCitationInfoLocal As MUSTER.Info.InspectionCitationInfo
        '    Try
        '        'Ownertable.Columns.Add("SELECTED1", GetType(Boolean))
        '        Ownertable.Columns.Add("OWNER_NAME", GetType(String))
        '        Ownertable.Columns.Add("OWNER_ID", GetType(Integer))

        '        'tbEntityTable.Columns.Add("SELECTED2", GetType(Boolean))
        '        tbEntityTable.Columns.Add("FCEID", GetType(Integer))
        '        tbEntityTable.Columns.Add("InspectionID", GetType(Integer))
        '        tbEntityTable.Columns.Add("OWNER_ID", GetType(Integer))
        '        tbEntityTable.Columns.Add("FacilityID", GetType(Integer))
        '        tbEntityTable.Columns.Add("FCEDate", GetType(Date))
        '        tbEntityTable.Columns.Add("Source", GetType(String))
        '        tbEntityTable.Columns.Add("DueDate", GetType(Date))
        '        tbEntityTable.Columns.Add("ReceivedDate", GetType(Date))
        '        tbEntityTable.Columns.Add("Owner", GetType(String))
        '        tbEntityTable.Columns.Add("Facility", GetType(String))
        '        tbEntityTable.Columns.Add("Inspector", GetType(String))
        '        tbEntityTable.Columns.Add("Inspected On", GetType(Date))
        '        tbEntityTable.Columns.Add("Citations", GetType(Integer))
        '        tbEntityTable.Columns.Add("CreatedBy", GetType(String))
        '        tbEntityTable.Columns.Add("CreatedOn", GetType(Date))
        '        tbEntityTable.Columns.Add("ModifiedBy", GetType(String))
        '        tbEntityTable.Columns.Add("ModifiedOn", GetType(Date))
        '        tbEntityTable.Columns.Add("Deleted", GetType(Boolean))

        '        CitationTable.Columns.Add("FCEID", GetType(Integer))
        '        CitationTable.Columns.Add("Citation", GetType(String))
        '        CitationTable.Columns.Add("Citation Text", GetType(String))
        '        CitationTable.Columns.Add("Category", GetType(String))
        '        CitationTable.Columns.Add("CCAT", GetType(String))
        '        CitationTable.Columns.Add("CCAT Comments", GetType(String))

        '        For Each oFCEInfoLocal In colFCE.Values
        '            dr = tbEntityTable.NewRow()
        '            'dr("SELECTED2") = oFCEInfoLocal.selected2
        '            dr("FCEID") = oFCEInfoLocal.ID
        '            dr("InspectionID") = oFCEInfoLocal.InspectionID
        '            dr("OWNER_ID") = oFCEInfoLocal.OwnerID
        '            dr("FacilityID") = oFCEInfoLocal.FacilityID
        '            If Date.Compare(oFCEInfoLocal.FCEDate, CDate("01/01/0001")) = 0 Then
        '                dr("FCEDate") = DBNull.Value
        '            Else
        '                dr("FCEDate") = oFCEInfoLocal.FCEDate
        '            End If

        '            dr("Source") = oFCEInfoLocal.Source
        '            If Date.Compare(oFCEInfoLocal.DueDate, CDate("01/01/0001")) = 0 Then
        '                dr("DueDate") = DBNull.Value
        '            Else
        '                dr("DueDate") = oFCEInfoLocal.DueDate
        '            End If
        '            If Date.Compare(oFCEInfoLocal.ReceivedDate, CDate("01/01/0001")) = 0 Then
        '                dr("ReceivedDate") = DBNull.Value
        '            Else
        '                dr("ReceivedDate") = oFCEInfoLocal.ReceivedDate
        '            End If
        '            dr("Facility") = oFCEInfoLocal.FacilityName
        '            dr("Inspector") = oFCEInfoLocal.InspectorName
        '            dr("Inspected On") = oFCEInfoLocal.InspectedOn
        '            dr("Citations") = oFCEInfoLocal.Citations
        '            dr("CreatedBy") = oFCEInfoLocal.CreatedBy
        '            dr("CreatedOn") = oFCEInfoLocal.CreatedOn
        '            dr("ModifiedBy") = oFCEInfoLocal.ModifiedBy
        '            dr("ModifiedOn") = oFCEInfoLocal.ModifiedOn
        '            dr("Deleted") = oFCEInfoLocal.Deleted
        '            tbEntityTable.Rows.Add(dr)

        '            Dim bolOwnerID As Boolean = False
        '            For Each tempRow As DataRow In Ownertable.Rows
        '                If tempRow.Item("OWNER_ID") = oFCEInfoLocal.OwnerID Then
        '                    bolOwnerID = True
        '                End If
        '            Next
        '            If Not bolOwnerID Then
        '                ownerRow = Ownertable.NewRow()
        '                'ownerRow("SELECTED") = oFCEInfoLocal.selected1
        '                ownerRow("OWNER_NAME") = oFCEInfoLocal.OwnerName
        '                ownerRow("OWNER_ID") = oFCEInfoLocal.OwnerID
        '                Ownertable.Rows.Add(ownerRow)
        '            End If
        '            For Each insCitationInfoLocal In pInsCitation.RetrieveByFCEID(oFCEInfoLocal.ID).Values
        '                CitationPenalty.Retrieve(insCitationInfoLocal.CitationID)
        '                CitationRow = CitationTable.NewRow()
        '                CitationRow("Citation") = CitationPenalty.StateCitation
        '                CitationRow("Citation Text") = CitationPenalty.Description
        '                CitationRow("Category") = CitationPenalty.Category
        '                CitationRow("CCAT") = insCitationInfoLocal.CCAT
        '                'code needs to be added for  CCAT Comments 
        '                CitationRow("CCAT Comments") = "TO DO"
        '                CitationRow("FCEID") = oFCEInfoLocal.ID
        '                CitationTable.Rows.Add(CitationRow)
        '            Next
        '            'CAEFceCitations.GetAll(oFCEInfoLocal.FacilityID)
        '            'For Each oFCECitationInfoLocal In CAEFceCitations.colFCECitations.Values
        '            '    oCitationInfoLocal = CitationPenalty.Retrieve(oFCECitationInfoLocal.CitationID)
        '            '    CitationRow("FacilityID") = oFCECitationInfoLocal.FacilityID
        '            'Next
        '        Next

        '        If Not colFCE.Count = 0 Then
        '            ds.Tables.Add(Ownertable)
        '            ds.Tables.Add(tbEntityTable)
        '            ds.Tables.Add(CitationTable)
        '            relation1 = New DataRelation("OwnertoFCE", ds.Tables(0).Columns("OWNER_ID"), ds.Tables(1).Columns("OWNER_ID"), False)
        '            relation2 = New DataRelation("FCEtoCitation", ds.Tables(1).Columns("FCEID"), ds.Tables(2).Columns("FCEID"), False)

        '            ds.Relations.Add(relation1)
        '            ds.Relations.Add(relation2)
        '            Return ds
        '        Else
        '            Return Nothing
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Function
#End Region
#End Region
#Region "External Event Handlers"
        Private Sub FacilityComplianceEventInfoChanged(ByVal bolValue As Boolean) Handles oFCEInfo.FCEInfoChanged
            RaiseEvent FacilityComplianceEventChanged(bolValue)
        End Sub
        Private Sub FacilityComplianceEventColChanged(ByVal bolValue As Boolean) Handles colFCE.FCEColChanged
            RaiseEvent ColChanged(bolValue)
        End Sub
#End Region
    End Class
End Namespace
