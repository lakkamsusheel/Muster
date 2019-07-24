'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.Entity
'   Provides the operations required to manipulate an Entity object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'   1.0         JVC2    11/19/04    Original class definition.
'   1.1         AN      01/03/05    Added Try catch and Exception Handling/Logging
'   1.2         AN      01/05/05    Added Events for Report Data Changed (Save/Cancel)
'   1.3         MNR     01/25/05    Added GetNext(), GetPrevious() and GetNextPrev(..) functions
'   1.4         JVC2    02/02/05    Added EntityTypeID to private members and initialize to "Report" type.
'   1.5         AB      02/22/05    Added DataAge check to the Retrieve function
'   1.6         JC      07/26/05    Added call to Archive method of info in SAVE to update object in memory.
'   1.7         JVC2    08/09/05    Added call to oAssociatedGroups.Reset on external call of Reset method.
'
' Function          Description
' GetEntity(NAME)   Returns the Entity requested by the string arg NAME
' GetEntity(ID)     Returns the Entity requested by the int arg ID
' GetEntityAll()    Returns an ReportsCollection with all Entity objects
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
' EntityCombo()     Returns a two-column datatable containing Name and ID for 
'                           the Entity objects in the internal ReportsCollection.
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pReport
#Region "Public Events"
        Public Event ReportExists(ByVal MsgStr As String)
        Public Event ReportChanged(ByVal BolValue As Boolean)
#End Region
#Region "Private Member Variables"
        Private colReports As Muster.Info.ReportsCollection
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private WithEvents oReportInfo As Muster.Info.ReportInfo
        Private oReportDB As New Muster.DataAccess.ReportDB
        Private oReportParams As New MUSTER.BusinessLogic.pReportParams
        'Private WithEvents oAssociatedGroups As New MUSTER.BusinessLogic.pProfile
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("Report").ID
#End Region
#Region "Constructors"
        Public Sub New()
            oReportInfo = New Muster.Info.ReportInfo
            colReports = New Muster.Info.ReportsCollection
            oReportParams = New MUSTER.BusinessLogic.pReportParams
            'oAssociatedGroups = New MUSTER.BusinessLogic.pProfile
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Integer
            Get
                Return oReportInfo.ID
            End Get

            Set(ByVal value As Integer)
                oReportInfo.ID = Integer.Parse(value)
            End Set
        End Property

        Public Property Name() As String
            Get
                Return oReportInfo.Name
            End Get

            Set(ByVal value As String)
                If Not colReports.Contains(value) Then
                    oReportInfo.Name = value
                Else
                    RaiseEvent ReportExists("Report " & value & " already exists!")
                End If
            End Set
        End Property

        Public Property Description() As String
            Get
                Return oReportInfo.Description
            End Get

            Set(ByVal value As String)
                oReportInfo.Description = value
            End Set
        End Property

        Public Property [Module]() As String
            Get
                Return oReportInfo.Module
            End Get

            Set(ByVal value As String)
                oReportInfo.Module = value
            End Set
        End Property

        Public Property Path() As String
            Get
                Return oReportInfo.Path
            End Get

            Set(ByVal value As String)
                oReportInfo.Path = value
            End Set
        End Property

        Public Property ReportParameter() As String
            Get
                Return oReportParams.Param
            End Get
            Set(ByVal Value As String)
                oReportParams.Param = Value
                'colParams(oParamInfo.ID) = oParamInfo
            End Set
        End Property

        Public Property ReportParameterDescription() As String
            Get
                If Not oReportParams Is Nothing Then
                    Return oReportParams.ParamDescription
                Else
                    Return String.Empty
                End If
            End Get

            Set(ByVal value As String)
                oReportParams.ParamDescription = value
                'colParams(oParamInfo.ID) = oParamInfo
            End Set
        End Property

        Public Property ReportParams() As MUSTER.BusinessLogic.pReportParams
            Get
                Return oReportParams
            End Get

            Set(ByVal value As MUSTER.BusinessLogic.pReportParams)
                oReportParams = value
            End Set
        End Property

        'Public Property UserGroups() As MUSTER.BusinessLogic.pProfile
        '    Get
        '        Return Me.oAssociatedGroups
        '    End Get

        '    Set(ByVal value As MUSTER.BusinessLogic.pProfile)
        '        Me.oAssociatedGroups = value
        '    End Set
        'End Property

        'Public ReadOnly Property EntityType() As Integer
        '    Get
        '        Return nEntityTypeID
        '    End Get
        'End Property

        Public Property Deleted() As Boolean
            Get
                Return oReportInfo.Deleted
            End Get

            Set(ByVal value As Boolean)
                oReportInfo.Deleted = value
            End Set
        End Property

        Public Property Active() As Boolean
            Get
                Return oReportInfo.Active
            End Get

            Set(ByVal value As Boolean)
                oReportInfo.Active = value
            End Set
        End Property

        Public Property IsDirty() As Boolean
            Get
                If oReportInfo.IsDirty Or oReportParams.colIsDirty Then
                    Return True
                Else
                    Return False
                End If
            End Get

            Set(ByVal value As Boolean)
                oReportInfo.IsDirty = value
            End Set
        End Property

        Public Property CreatedBy() As String
            Get
                Return oReportInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oReportInfo.CreatedBy = Value
            End Set
        End Property

        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oReportInfo.CreatedOn
            End Get
        End Property

        Public Property ModifiedBy() As String
            Get
                Return oReportInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oReportInfo.ModifiedBy = Value
            End Set
        End Property

        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oReportInfo.ModifiedOn
            End Get
        End Property

        Public Property ReportGroupRelationCollection() As MUSTER.Info.ReportGroupRelationsCollection
            Get
                Return oReportInfo.ReportGroupRelationCollection
            End Get
            Set(ByVal Value As MUSTER.Info.ReportGroupRelationsCollection)
                oReportInfo.ReportGroupRelationCollection = Value
            End Set
        End Property

#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        ' Obtains and returns an entity as called for by name
        Public Function Retrieve(ByVal Name As String) As MUSTER.Info.ReportInfo
            Try
                If colReports.Contains(Name) Then
                    oReportInfo = colReports.Item(Name)
                    If oReportInfo.IsAgedData = True And oReportInfo.IsDirty = False Then
                        colReports.Remove(oReportInfo)
                        Return Retrieve(Name)
                    Else
                        'oReportParams.Clear()
                        'oReportParams.Retrieve("SYSTEM|REPORTPARAMS|" & oReportInfo.filename, False)
                        oReportParams.ReportID = oReportInfo.ID
                        'oAssociatedGroups.Clear()
                        'oAssociatedGroups.Retrieve(oReportInfo.Name & "|USER GROUPS")
                        oReportInfo.ReportGroupRelationCollection = oReportDB.DBGetReportGroupRel(oReportInfo.ID)
                        Return oReportInfo
                    End If
                Else
                    'oAssociatedGroups.Clear()
                    'oAssociatedGroups.Retrieve(Name & "|USER GROUPS")
                    colReports.Add(oReportDB.DBGetByName(Name))
                    Return Retrieve(Name)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        'Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal ID As Int64) As MUSTER.Info.ReportInfo
            Dim oReportInfoLocal As MUSTER.Info.ReportInfo
            Dim bolDataAged As Boolean = False

            Try
                For Each oReportInfoLocal In colReports.Values
                    If oReportInfoLocal.ID = ID Then
                        If oReportInfoLocal.IsAgedData = True And oReportInfoLocal.IsDirty = False Then
                            bolDataAged = True
                            Exit For
                        Else
                            oReportInfo = oReportInfoLocal
                            'oReportParams.Clear()
                            'oReportParams.Retrieve("SYSTEM|REPORTPARAMS|" & oReportInfoLocal.FileName, False)
                            'oAssociatedGroups.Clear()
                            'oAssociatedGroups.Retrieve(oReportInfo.Name & "|USER GROUPS")
                            oReportInfo.ReportGroupRelationCollection = oReportDB.DBGetReportGroupRel(oReportInfo.ID)
                            oReportParams.ReportID = ID
                            Return oReportInfo
                        End If
                    End If
                Next

                If bolDataAged Then
                    colReports.Remove(oReportInfoLocal)
                End If

                oReportInfo = oReportDB.DBGetByID(ID)
                'oAssociatedGroups.Clear()
                'oAssociatedGroups.Retrieve(oReportInfo.Name & "|USER GROUPS")
                oReportInfo.ReportGroupRelationCollection = oReportDB.DBGetReportGroupRel(oReportInfo.ID)
                colReports.Add(oReportInfo)
                Return oReportInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        'Saves the data in the current Info object
        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try
                Dim oldReportID As Integer = oReportInfo.ID
                oReportDB.Put(oReportInfo, moduleID, staffID, returnVal)
                If Not returnVal = String.Empty Then
                    Exit Sub
                End If
                If oldReportID <> oReportInfo.ID Then
                    Dim oldRGIDs As New Collection
                    Dim reportGroupRelInfo As MUSTER.Info.ReportGroupRelationInfo
                    For Each reportGroupRelInfo In oReportInfo.ReportGroupRelationCollection.Values
                        If reportGroupRelInfo.IsDirty Then
                            oldRGIDs.Add(reportGroupRelInfo.ID)
                        End If
                    Next
                    If Not oldRGIDs Is Nothing Then
                        For index As Integer = 1 To oldRGIDs.Count
                            Dim colKey As String = CType(oldRGIDs.Item(index), String)
                            reportGroupRelInfo = oReportInfo.ReportGroupRelationCollection.Item(colKey)
                            reportGroupRelInfo.ReportID = oReportInfo.ID
                            oReportInfo.ReportGroupRelationCollection.ChangeKey(colKey, reportGroupRelInfo.ReportID.ToString + "|" + reportGroupRelInfo.GroupID.ToString)
                        Next
                    End If
                End If
                FlushReportGroupRel(moduleID, staffID, returnVal)
                If Not returnVal = String.Empty Then
                    Exit Sub
                End If
                oReportParams.Flush(moduleID, staffID, returnVal)
                If Not returnVal = String.Empty Then
                    Exit Sub
                End If
                'oAssociatedGroups.Flush(moduleID, staffID, returnVal)
                If Not returnVal = String.Empty Then
                    Exit Sub
                End If
                oReportInfo.Archive()
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Private Sub FlushReportGroupRel(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            For Each reportGroupRelInfo As MUSTER.Info.ReportGroupRelationInfo In oReportInfo.ReportGroupRelationCollection.Values
                If reportGroupRelInfo.IsDirty And reportGroupRelInfo.ReportID > 0 Then
                    oReportDB.PutReportGroupRel(reportGroupRelInfo, moduleID, staffID, returnVal)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                    reportGroupRelInfo.Archive()
                End If
            Next
        End Sub
#End Region
#Region "Collection Operations"
        Function GetAll() As MUSTER.Info.ReportsCollection
            Try
                colReports.Clear()
                colReports = oReportDB.GetAllInfo
                Return colReports
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal ID As Int64)
            Try
                oReportInfo = oReportDB.DBGetByID(ID)
                oReportParams.Clear()
                oReportParams.Retrieve("SYSTEM|REPORTPARAMS|" & oReportInfo.FileName, False)
                'oAssociatedGroups.Clear()
                'oAssociatedGroups.Retrieve(oReportInfo.Name & "|USER GROUPS")
                oReportInfo.ReportGroupRelationCollection = oReportDB.DBGetReportGroupRel(oReportInfo.ID)
                colReports.Add(oReportInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        'Adds an entity to the collection as called for by Name
        Public Sub Add(ByVal Name As String)
            Try
                oReportInfo = oReportDB.DBGetByName(Name)
                oReportParams.Clear()
                oReportParams.Retrieve("SYSTEM|REPORTPARAMS|" & oReportInfo.FileName, False)
                'oAssociatedGroups.Clear()
                'oAssociatedGroups.Retrieve(oReportInfo.Name & "|USER GROUPS")
                oReportInfo.ReportGroupRelationCollection = oReportDB.DBGetReportGroupRel(oReportInfo.ID)
                colReports.Add(oReportInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oReport As MUSTER.Info.ReportInfo)

            Try
                oReportInfo = oReport
                oReportParams.Clear()
                oReportParams.Retrieve("SYSTEM|REPORTPARAMS|" & oReport.FileName, False)
                'oAssociatedGroups.Clear()
                'oAssociatedGroups.Retrieve(oReportInfo.Name & "|USER GROUPS")
                oReportInfo.ReportGroupRelationCollection = oReportDB.DBGetReportGroupRel(oReportInfo.ID)
                colReports.Add(oReportInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Int64)

            Dim myIndex As Int16 = 1
            Dim oRptInf As MUSTER.Info.ReportInfo

            Try
                For Each oRptInf In colReports.Values
                    If oRptInf.ID = ID Then
                        colReports.Remove(oRptInf)
                        Exit Sub
                    End If
                    myIndex += 1
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

            'Throw New Exception("Report " & ID.ToString & " is not in the collection of reports.")

        End Sub
        'Removes the entity called for by Name from the collection
        Public Sub Remove(ByVal Name As String)

            Try
                If colReports.Contains(Name) Then
                    colReports.Remove(Name)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

            'Throw New Exception("Report " & Name & " is not in the collection of reports.")

        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oRptInf As MUSTER.Info.ReportInfo)

            Try
                If colReports.Contains(oRptInf.Name) Then
                    colReports.Remove(oRptInf)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

            'Throw New Exception("Report " & oRptInf.Name & " is not in the collection of reports.")

        End Sub
        'Public Sub AddToUserGroup(ByVal strName As String)
        '    Try
        '        Dim oProfileInfo As Muster.Info.ProfileInfo
        '        oProfileInfo = New Muster.Info.ProfileInfo(Me.Name, "USER GROUPS", strName, "NONE", "NONE", False, "", Now(), "", Now())
        '        oProfileInfo.ProfileValue = strName
        '        oProfileInfo.IsDirty = True
        '        oAssociatedGroups.Add(oProfileInfo)
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Sub
        '' Remove the user from a user group
        'Public Sub RemoveFromUserGroup(ByVal strName As String)
        '    Try
        '        Dim strKey As String = ""
        '        Dim test As Muster.Info.ProfileInfo
        '        strKey = Me.Name & "|USER GROUPS|" & strName & "|NONE"
        '        oAssociatedGroups.Retrieve(strKey, True).IsDirty = True
        '        oAssociatedGroups.ProfileValue = "NONE"
        '        'oAssociatedGroups.colIsDirty = True
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Sub
        Public Function GetNext() As String
            Return GetNextPrev(1)
        End Function
        Public Function GetPrevious() As String
            Return GetNextPrev(-1)
        End Function
        Private Function GetNextPrev(ByVal direction As Integer) As String
            Dim strArr() As String = colReports.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colReports.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            Else
                Return colReports.Item(nArr.GetValue(colIndex)).ID.ToString
            End If
        End Function
#End Region
#Region "General Operations"
        'Public Function ListReportNames(ByVal bolshowdeleted As Boolean) As DataTable
        '    '6/7/2005 - AN - REMOVED From BLL and ADDED TO DAL
        '    'Dim dtReportNames As DataTable

        '    'Dim strSQL As String
        '    'Dim dsset As New DataSet
        '    'strSQL = "SELECT '' as REPORT_NAME,'' as Report_ID UNION SELECT  REPORT_NAME,REPORT_ID FROM tblSYS_REPORT_MASTER "
        '    'strSQL += IIf(Not bolshowdeleted, " WHERE ACTIVE = 1", "")
        '    'strSQL += " Order by REPORT_NAME"
        '    'Try
        '    '    dsset = oReportDB.DBGetDS(strSQL)
        '    '    If dsset.Tables(0).Rows.Count > 0 Then
        '    '        dtReportNames = dsset.Tables(0)
        '    '    Else
        '    '        dtReportNames = Nothing
        '    '    End If
        '    '    Return dtReportNames
        '    Try
        '        Return oReportDB.ListReportNames(bolshowdeleted)
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Function
        'Public Function ListReportNames(ByVal moduleid As String, ByVal bolshowdeleted As Boolean, Optional ByVal UserID As String = "") As DataTable
        '    '6/7/2005 - AN - REMOVED From BLL and ADDED TO DAL
        '    'Dim dtReportNames As DataTable

        '    'Dim strSQL As String
        '    'Dim dsset As New DataSet
        '    'strSQL = "SELECT '' as REPORT_NAME,'' as Report_ID,'' as REPORT_LOC UNION SELECT  REPORT_NAME as REPORT_NAME,report_id,REPORT_LOC as Report_ID FROM tblSYS_REPORT_MASTER "
        '    'strSQL += IIf(moduleid > 0, " Where REPORT_MODULE='" & moduleid & "'", "")
        '    'strSQL += IIf(Not bolshowdeleted, IIf(moduleid > 0, " and ACTIVE = 1", " where ACTIVE = 1"), "")
        '    'strSQL += IIf(UserID > 0, "Report_NAME in (Select USER_ID from tblSys_Profile_Info where USER_ID in (SELECT  REPORT_NAME FROM tblSYS_REPORT_MASTER) and PROFILE_KEY='USER GROUPS' and PROFILE_VALUE in (Select PROFILE_VALUE from tblSys_Profile_Info Where USER_ID='" & UserID & "' and PROFILE_KEY='USER GROUPS'))", "")
        '    'strSQL += " Order by REPORT_NAME"
        '    'Try
        '    '    dsset = oReportDB.DBGetDS(strSQL)
        '    '    If dsset.Tables(0).Rows.Count > 0 Then
        '    '        dtReportNames = dsset.Tables(0)
        '    '    Else
        '    '        dtReportNames = Nothing
        '    '    End If
        '    '    Return dtReportNames
        '    Try
        '        Return oReportDB.ListReportNames(moduleid, bolshowdeleted, UserID)
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Function

        Public Sub Clear()
            oReportInfo = New MUSTER.Info.ReportInfo
            oReportParams = New MUSTER.BusinessLogic.pReportParams
        End Sub
        Public Sub Reset()
            oReportInfo.Reset()
            'oAssociatedGroups.Reset()
            If oReportParams.ReportID > 0 Then
                oReportParams.Reset()
            End If
            RaiseEvent ReportChanged(oReportInfo.IsDirty)
        End Sub
#End Region
#Region "Miscellaneous Operations"
        'Returns a datatable of the entities in the collection
        Public Function EntityTable() As DataTable

            Dim oReportInfoLocal As MUSTER.Info.ReportInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable

            Try
                tbEntityTable.Columns.Add("Report ID")
                tbEntityTable.Columns.Add("Report Name")
                tbEntityTable.Columns.Add("Report Description")
                tbEntityTable.Columns.Add("Report Location")
                tbEntityTable.Columns.Add("Report Module")
                tbEntityTable.Columns.Add("Active")
                tbEntityTable.Columns.Add("Created By")
                tbEntityTable.Columns.Add("Created On")
                tbEntityTable.Columns.Add("Modified By")
                tbEntityTable.Columns.Add("Modified On")

                For Each oReportInfoLocal In colReports.Values
                    dr = tbEntityTable.NewRow()
                    dr("Entity ID") = oReportInfoLocal.ID
                    dr("Entity Name") = oReportInfoLocal.Name
                    dr("Report Description") = oReportInfoLocal.Description
                    dr("Report Location") = oReportInfoLocal.Path
                    dr("Report Module") = oReportInfoLocal.Module
                    dr("Active") = oReportInfoLocal.Deleted
                    dr("Created By") = oReportInfoLocal.CreatedBy
                    dr("Created On") = oReportInfoLocal.CreatedOn
                    dr("Modified By") = oReportInfoLocal.ModifiedBy
                    dr("Modified On") = oReportInfoLocal.ModifiedOn
                    tbEntityTable.Rows.Add(dr)
                Next
                Return tbEntityTable
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        'Returns a two-column datatable of the entities in the collection column names Entity ID and Entity Name
        Public Function ReportCombo() As DataTable

            Dim oReportInfoLocal As MUSTER.Info.ReportInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable

            Try
                tbEntityTable.Columns.Add("Report ID")
                tbEntityTable.Columns.Add("Report Name")

                'For Each oReportInfo In colReports
                For Each oReportInfoLocal In colReports.Values
                    dr = tbEntityTable.NewRow()
                    dr("Report ID") = oReportInfoLocal.ID
                    dr("Report Name") = oReportInfoLocal.Name
                    tbEntityTable.Rows.Add(dr)
                Next
                Return tbEntityTable
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function

        Public Function ReportFileExists(ByVal FilePath As String, ByVal CurrentReportID As Integer) As Boolean
            Dim oRptInf As MUSTER.Info.ReportInfo
            Try
                Return oReportDB.ReportFileExists(FilePath, CurrentReportID)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function GetReportsForUser(Optional ByVal staffID As Integer = 0, Optional ByVal moduleID As Integer = 0, Optional ByVal showDeleted As Boolean = False, Optional ByVal showFavReport As Boolean = False) As DataTable
            Try
                Return oReportDB.DBGetReportsForUser(staffID, moduleID, showDeleted, showFavReport).Tables(0)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Sub SaveFavReport(ByVal staffID As Integer, ByVal reportID As Integer, ByVal deleted As Boolean, ByVal moduleID As Integer, ByVal securityStaffID As Integer, ByRef returnVal As String)
            oReportDB.PutFavReport(staffID, reportID, deleted, moduleID, securityStaffID, returnVal)
        End Sub
        Public Function GetCapPreMonthly(Optional ByVal showPrev As Boolean = False) As DataSet
            Try
                Dim dsReturn As DataSet = oReportDB.DBGetCapPreMonthly(showPrev)
                Dim dsRel1, dsRel2 As DataRelation
                If dsReturn.Tables.Count > 0 Then
                    dsRel1 = New DataRelation("Tank", dsReturn.Tables(0).Columns("CAP_STATUS_ID"), dsReturn.Tables(1).Columns("CAP_STATUS_ID"))
                    dsRel2 = New DataRelation("Pipe", dsReturn.Tables(0).Columns("CAP_STATUS_ID"), dsReturn.Tables(2).Columns("CAP_STATUS_ID"))
                    dsReturn.Relations.Add(dsRel1)
                    dsReturn.Relations.Add(dsRel2)
                End If
                Return dsReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Sub SaveCapPreMonthly(ByVal strCapStatusIDs As String)
            oReportDB.PutCapPreMonthly(strCapStatusIDs)
        End Sub
#End Region
#End Region
#Region "External Event Handlers"
        Private Sub ThisReport(ByVal bolValue As Boolean) Handles oReportInfo.ReportChanged
            '
            ' Alert the client that the current report data has changed
            '
            RaiseEvent ReportChanged(Me.IsDirty Or oReportParams.colIsDirty)
        End Sub
#End Region
    End Class
End Namespace
