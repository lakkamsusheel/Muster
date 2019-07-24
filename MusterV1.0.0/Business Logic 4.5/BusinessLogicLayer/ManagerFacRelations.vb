'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.ManagerFacRelations
'   Provides the operations required to manipulate an Entity object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'                                Original class definition
'         Hua Cao     09/15/12    Added and Modified Functions and Attributes
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
' NOTE: This file to be used as ManagerFacRelations to build other objects.
'       Replace keyword "ManagerFacRelations" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pManagerFacRelations
#Region "Public Events"
        Public Event ManagerFacRelationsErr(ByVal MsgStr As String)
        Public Event ManagerFacRelationsChanged(ByVal bolValue As Boolean)
        Public Event ColChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("ManagerFacRelations").ID
        Private WithEvents oManagerFacRelationsInfo As MUSTER.Info.ManagerFacRelationInfo
        Private WithEvents colManagerFacRelations As MUSTER.Info.ManagerFacRelationCollection
        Private oManagerFacRelationsDB As New MUSTER.DataAccess.RelationDB
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private nID As Int64 = -1
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
#End Region
#Region "Constructors"
        Public Sub New()
            oManagerFacRelationsInfo = New MUSTER.Info.ManagerFacRelationInfo
            colManagerFacRelations = New MUSTER.Info.ManagerFacRelationCollection
        End Sub
        '********************************************************
        '
        ' Overloaded NEW which will populate with a single instance
        '   of the named ManagerFacRelations object.
        '
        '********************************************************
        Public Sub New(ByVal ManagerFacRelationsName As String)
            oManagerFacRelationsInfo = New MUSTER.Info.ManagerFacRelationInfo
            colManagerFacRelations = New MUSTER.Info.ManagerFacRelationCollection
            Me.Retrieve(ManagerFacRelationsName)
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Integer
            Get
                Return oManagerFacRelationsInfo.ID
            End Get
            Set(ByVal Value As Integer)
                oManagerFacRelationsInfo.ID = Integer.Parse(Value)
            End Set
        End Property
        Public Property ManagerID() As Integer
            Get
                Return oManagerFacRelationsInfo.ManagerID
            End Get
            Set(ByVal Value As Integer)
                oManagerFacRelationsInfo.ManagerID = Integer.Parse(Value)
            End Set
        End Property
        Public Property FacilityID() As Integer
            Get
                Return oManagerFacRelationsInfo.FacilityID
            End Get
            Set(ByVal Value As Integer)
                oManagerFacRelationsInfo.FacilityID = Value
            End Set
        End Property

        Public Property RelationID() As Integer
            Get
                Return oManagerFacRelationsInfo.RelationID
            End Get
            Set(ByVal Value As Integer)
                oManagerFacRelationsInfo.RelationID = Value
            End Set
        End Property


        Public Property RelationDesc() As String
            Get
                Return oManagerFacRelationsInfo.RelationDesc
            End Get
            Set(ByVal Value As String)
                oManagerFacRelationsInfo.RelationDesc = Value
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oManagerFacRelationsInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oManagerFacRelationsInfo.Deleted = Boolean.Parse(Value)
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oManagerFacRelationsInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oManagerFacRelationsInfo.IsDirty = Boolean.Parse(value)
            End Set
        End Property
        Public Property colIsDirty() As Boolean
            Get
                Dim xManagerFacRelationsinfo As MUSTER.Info.ManagerFacRelationInfo
                For Each xManagerFacRelationsinfo In colManagerFacRelations.Values
                    If xManagerFacRelationsinfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
            Set(ByVal Value As Boolean)
                oManagerFacRelationsInfo.IsDirty = Value
            End Set
        End Property
        Public Property MgrFacRelationInfo() As MUSTER.Info.ManagerFacRelationInfo
            Get
                Return Me.MgrFacRelationInfo
            End Get

            Set(ByVal value As MUSTER.Info.ManagerFacRelationInfo)
                Me.MgrFacRelationInfo = value
            End Set
        End Property
        Public Property colMgrFacRelation() As MUSTER.Info.ManagerFacRelationCollection
            Get
                Return Me.colManagerFacRelations
            End Get

            Set(ByVal value As MUSTER.Info.ManagerFacRelationCollection)
                Me.colManagerFacRelations = value
            End Set
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        'Obtains and returns an entity as called for by ID
        'Public Function Retrieve(ByVal ManagerID As Integer) As MUSTER.Info.ManagerFacRelationInfo
        '    Dim oManagerFacRelationsInfoLocal As MUSTER.Info.ManagerFacRelationInfo
        '    Try
        '        For Each oManagerFacRelationsInfoLocal In colManagerFacRelations.Values
        '            If oManagerFacRelationsInfoLocal.ManagerID = ManagerID Then
        '                oManagerFacRelationsInfo = oManagerFacRelationsInfoLocal
        '                Return oManagerFacRelationsInfo
        '            End If
        '        Next
        '        oManagerFacRelationsInfo = oManagerFacRelationsDB.DBGetByManagerID(ManagerID)
        '        If oManagerFacRelationsInfo.ManagerID = 0 Then
        '            oManagerFacRelationsInfo.ManagerID = nID
        '            nID -= 1
        '        End If
        '        colManagerFacRelations.Add(oManagerFacRelationsInfo)
        '        Return oManagerFacRelationsInfo
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Function
        Public Function Retrieve(ByVal ID As Integer) As MUSTER.Info.ManagerFacRelationInfo
            Try
                '       oManagerFacRelationsInfo = Nothing
                Dim oManagerFacRelationsInfoLocal As MUSTER.Info.ManagerFacRelationInfo

                Try
                    For Each oManagerFacRelationsInfoLocal In colManagerFacRelations.Values
                        If oManagerFacRelationsInfoLocal.ID = ID Then
                            oManagerFacRelationsInfo = oManagerFacRelationsInfoLocal
                            Return oManagerFacRelationsInfo
                        End If
                    Next
                    oManagerFacRelationsInfo = oManagerFacRelationsDB.DBGetByID(ID)
                    If oManagerFacRelationsInfo.ID = 0 Then
                        oManagerFacRelationsInfo.ID = nID
                        nID -= 1
                    End If
                    colManagerFacRelations.Add(oManagerFacRelationsInfo)
                    Return oManagerFacRelationsInfo
                Catch ex As Exception
                    If InStr(UCase(ex.source), UCase("muster.businesslogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                    Throw ex
                End Try


                'If colManagerFacRelations.Contains(ManagerFacRelationsName) Then
                '    oManagerFacRelationsInfo = colManagerFacRelations(ManagerFacRelationsName)
                'Else
                '    If oManagerFacRelationsInfo Is Nothing Then
                '        oManagerFacRelationsInfo = New MUSTER.Info.ManagerFacRelationInfo
                '    End If
                '    colManagerFacRelations.Add(oManagerFacRelationsInfo)
                'End If
                'Return oManagerFacRelationsInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        Public Function Retrieve(ByVal ManagerFacRelationName As String) As MUSTER.Info.ManagerFacRelationInfo
            Try
                oManagerFacRelationsInfo = Nothing
                If colManagerFacRelations.Contains(ManagerFacRelationName) Then
                    oManagerFacRelationsInfo = colManagerFacRelations(ManagerFacRelationName)
                Else
                    If oManagerFacRelationsInfo Is Nothing Then
                        oManagerFacRelationsInfo = New MUSTER.Info.ManagerFacRelationInfo
                    End If
                    colManagerFacRelations.Add(oManagerFacRelationsInfo)
                End If
                Return oManagerFacRelationsInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        'Saves the data in the current Info object
        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal bolValidated As Boolean = False)

            Try
                'If Me.ValidateData() Then
                Dim OldKey As String = oManagerFacRelationsInfo.ID.ToString
                oManagerFacRelationsDB.Put(oManagerFacRelationsInfo, moduleID, staffID, returnVal)
                If Not returnVal = String.Empty Then
                    Exit Sub
                End If

                If Not bolValidated Then
                    If oManagerFacRelationsInfo.ID.ToString <> OldKey Then
                        colMgrFacRelation.ChangeKey(OldKey, oManagerFacRelationsInfo.ID.ToString)
                    End If
                End If
                oManagerFacRelationsInfo.Archive()
                oManagerFacRelationsInfo.IsDirty = False
                'End If

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Validates the data before saving
        Public Function ValidateData() As Boolean
            Dim errStr As String = ""
            Dim validateSuccess As Boolean = False

            Try

                Return validateSuccess
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region
#Region "Collection Operations"
        ''Gets all the info
        Function GetAll(Optional ByVal nManagerID As Integer = 0) As MUSTER.Info.ManagerFacRelationCollection
            Try
                ' colManagerFacRelations.Clear()
                colManagerFacRelations = oManagerFacRelationsDB.DBGetByManagerID(nManagerID)
                Return colManagerFacRelations
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal ID As Integer)
            Try
                Dim oMgrRelationInfo As MUSTER.Info.ManagerFacRelationInfo
                If ID = 0 Then
                    oMgrRelationInfo = New MUSTER.Info.ManagerFacRelationInfo
                    oMgrRelationInfo.ID = ID
                    nID -= 1
                    oManagerFacRelationsInfo = oMgrRelationInfo
                Else
                    oManagerFacRelationsInfo = oManagerFacRelationsDB.DBGetByID(ID)
                End If

                colManagerFacRelations.Add(oManagerFacRelationsInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oManagerFacRelations As MUSTER.Info.ManagerFacRelationInfo)
            Try
                oManagerFacRelationsInfo = oManagerFacRelations
                colManagerFacRelations.Add(oManagerFacRelationsInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Integer)
            Dim myIndex As Int16 = 1
            Dim oManagerFacRelationsInfoLocal As MUSTER.Info.ManagerFacRelationInfo

            Try
                For Each oManagerFacRelationsInfoLocal In colManagerFacRelations.Values
                    If oManagerFacRelationsInfoLocal.ID = ID Then
                        colManagerFacRelations.Remove(oManagerFacRelationsInfoLocal)
                        Exit Sub
                    End If
                    myIndex += 1
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("ManagerFacRelations " & ID.ToString & " is not in the collection of ManagerFacRelations.")
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oManagerFacRelations As MUSTER.Info.ManagerFacRelationInfo)
            Try
                colManagerFacRelations.Remove(oManagerFacRelations)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("ManagerFacRelations " & oManagerFacRelations.ID & " is not in the collection of ManagerFacRelations.")
        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal nLicenseeID As Integer = 0)

            Dim IDs As New Collection
            Dim index As Integer
            Dim xManagerFacRelationsInfo As MUSTER.Info.ManagerFacRelationInfo

            For Each xManagerFacRelationsInfo In colManagerFacRelations.Values
                If xManagerFacRelationsInfo.IsDirty Then
                    If nLicenseeID > 0 Then
                        xManagerFacRelationsInfo.ManagerID = nLicenseeID
                    End If
                    oManagerFacRelationsInfo = xManagerFacRelationsInfo
                    IDs.Add(oManagerFacRelationsInfo.ID)
                    Me.Save(moduleID, staffID, returnVal, True)
                End If
            Next

            If Not (IDs Is Nothing) Then
                For index = 1 To IDs.Count
                    Dim colKey As String = CType(IDs.Item(index), String)
                    xManagerFacRelationsInfo = colManagerFacRelations.Item(colKey)
                    colManagerFacRelations.ChangeKey(colKey, xManagerFacRelationsInfo.ID)
                Next
            End If

        End Sub
        Public Function GetNext() As String
            Return GetNextPrev(1)
        End Function
        Public Function GetPrevious() As String
            Return GetNextPrev(-1)
        End Function
        Private Function GetNextPrev(ByVal direction As Integer) As String
            Dim strArr() As String = colManagerFacRelations.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colManagerFacRelations.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            Else
                Return colManagerFacRelations.Item(nArr.GetValue(colIndex)).ID.ToString
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oManagerFacRelationsInfo = New MUSTER.Info.ManagerFacRelationInfo
        End Sub
        Public Sub Reset()
            oManagerFacRelationsInfo.Reset()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function RelationTable() As DataTable
            Dim oManagerFacRelationsInfoLocal As New MUSTER.Info.ManagerFacRelationInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Try
                tbEntityTable.Columns.Add("MGRFACRELATION_ID", GetType(Integer))
                tbEntityTable.Columns.Add("FacilityID", GetType(Integer))
                tbEntityTable.Columns.Add("ManagerID", GetType(Integer))
                tbEntityTable.Columns.Add("Relation", GetType(Integer))
                tbEntityTable.Columns.Add("Deleted", GetType(Boolean))
                tbEntityTable.Columns("Deleted").DefaultValue = False

                tbEntityTable.Columns.Add("RelationDesc")
                '  If Not colManagerFacRelations Is Nothing Then
                For Each oManagerFacRelationsInfoLocal In colManagerFacRelations.Values
                    dr = tbEntityTable.NewRow()
                    dr("MGRFACRELATION_ID") = oManagerFacRelationsInfoLocal.ID
                    dr("FacilityID") = oManagerFacRelationsInfoLocal.FacilityID
                    dr("ManagerID") = oManagerFacRelationsInfoLocal.ManagerID
                    dr("Relation") = oManagerFacRelationsInfoLocal.RelationID
                    dr("RelationDesc") = oManagerFacRelationsInfoLocal.RelationDesc
                    dr("Deleted") = oManagerFacRelationsInfoLocal.Deleted
                    tbEntityTable.Rows.Add(dr)
                Next
                '  Else
                '  dr = tbEntityTable.NewRow()
                '  tbEntityTable.Rows.Add(dr)

                '  End If
                Return tbEntityTable
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region
#Region "LookUp Operations"
        Public Function ListCourseTypes(Optional ByVal showBlankPropertyName As Boolean = True) As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vCOM_COURSETYPE")
                If showBlankPropertyName Then
                    Dim dr As DataRow = dtReturn.NewRow
                    For Each dtCol As DataColumn In dtReturn.Columns
                        If dtCol.DataType.Name.IndexOf("String") > -1 Then
                            dr(dtCol) = " "
                        ElseIf dtCol.DataType.Name.IndexOf("Int") > -1 Then
                            dr(dtCol) = 0
                        End If
                    Next
                    dtReturn.Rows.InsertAt(dr, 0)
                End If
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function ListRelations() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vCOM_Relation")
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function ListProviders() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vCOM_PROVIDER")
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Private Function GetDataTable(ByVal DBViewName As String) As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String
            Try
                strSQL = "SELECT * FROM " & DBViewName

                dsReturn = oManagerFacRelationsDB.DBGetDS(strSQL)
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

#End Region
#Region "External Event Handlers"
        Private Sub PManagerFacRelationsInfoChanged(ByVal bolValue As Boolean) Handles oManagerFacRelationsInfo.MgrFacRelationInfoChanged
            RaiseEvent ManagerFacRelationsChanged(bolValue)
        End Sub
        Private Sub PManagerFacRelationsColChanged(ByVal bolValue As Boolean) Handles colManagerFacRelations.ManagerFacRelationColChanged
            RaiseEvent ColChanged(bolValue)
        End Sub
#End Region
    End Class
End Namespace
