'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.Profile
'   Provides the collection and info object to the client for manipulating
'     ProfileInfo data.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'   1.0         JVC2    11/19/04    Original class definition.
'   1.1         AN      12/30/04    Added Try catch and Exception Handling/Logging
'   1.2         MNR     01/25/05    Added GetNext(), GetPrevious() and GetNextPrev(..) functions
'   1.3         JVC2    02/02/05    Added EntityTypeID to private members and initialize to "Profile" type.
'                                       Also added EntityType attribute to expose the typeID.
'   1.4         JVC2    02/08/2005  Added ProfileUserTable function.
'   1.5         MR      02/15/2005  Altered Query in ProfileUserTable and ProfileKeyTable.
'   1.6         AB      02/22/05    Added DataAge check to the Retrieve function
'   1.7         JVC2    02/22/05    Modified Retrieve to return either the first of a group of profile
'                                       entries or the one that was retrieved if only one match was found
'                                       on load from the data repostitory.
'   1.8         JVC2    02/24/05    Modified Retrieve to return the info object directly after adding to
'                                       the collection if there is only one object in the collection.  This
'                                       is due to the fact that if a partial key is specified by the caller,
'                                       the function recurses infinitely attempting to find a match that will
'                                       never be present.
' Operations
' Function                      Description
' New()                         Initializes the ProfileCollection and ProfileInfo objects.
' Retrieve(Key, [ShowDeleted])  Sets the internal ProfileInfo to the ProfileInfo matching the 
'                                supplied key.  In the event that a partial key is supplied,
'                                all matching ProfileInfos are populated to the internal 
'                                ProfileCollection and the internal ProfileInfo object is
'                                set to the first member of the collection.  In either case,
'                                the internal ProfileInfo object is returned to the client.
' Save()                        Marshalls the internal ProfileInfo to the repository.
' Add(ProfileInfo)              Adds the supplied ProfileInfo object to the internal ProfileCollection
'                                and sets the internal ProfileInfo object to same.
' Remove(ProfileInfo)           Removes the supplied ProfileInfo object from the internal ProfileCollection
'                                if it is contained by the collection.
' Items()                       Returns the internal ProfileCollection (a Dictionary object) to the client.
' Values()                      Returns the collection of ProfileInfo objects to the client (used in for..next).
' Flush()                       Marshalls all modified/new ProfileInfo objects in the ProfileCollection
'                                to the repository.
' Clear()                       Sets the internal ProfileInfo object to an empty ProfileInfo and clears
'                                the internal ProfileCollection.
' Reset()                       Reverts the internal ProfileInfo object to it's state when last retrieved
'                                from or marshalled to the repository.
' GetTable(strSQL)              Returns a datatable containing the resultset from the supplied SQL string.
' ProfileTable()                Returns a datatable containing rows corresponding to the ProfileInfo
'                                objects in the internal ProfileCollection.
'-------------------------------------------------------------------------------
'
' TODO - Integrate with solution 2/9/2005 - JVC2

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pProfile
#Region "Private Member Variables"
        Private WithEvents colProfiles As Muster.Info.ProfileCollection
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private WithEvents oProfileInfo As Muster.Info.ProfileInfo
        Private oProfileDB As New Muster.DataAccess.ProfileDB
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("Profile").ID
#End Region
#Region "Public Events"
        Public Event InfoBecameDirty(ByVal BolValue As Boolean)
#End Region
#Region "Constructors"
        Public Sub New()
            oProfileInfo = New Muster.Info.ProfileInfo
            colProfiles = New Muster.Info.ProfileCollection
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As String
            Get
                Return oProfileInfo.ID
            End Get

            Set(ByVal value As String)
                oProfileInfo.ID = value
            End Set
        End Property

        Public Property ProfileKey() As String
            Get
                Return oProfileInfo.ProfileKey
            End Get
            Set(ByVal Value As String)
                oProfileInfo.ProfileKey = Value
                colProfiles(oProfileInfo.ID) = oProfileInfo
            End Set
        End Property

        Public Property ProfileMod1() As String
            Get
                Return oProfileInfo.ProfileMod1
            End Get
            Set(ByVal Value As String)
                oProfileInfo.ProfileMod1 = Value
                colProfiles(oProfileInfo.ID) = oProfileInfo
            End Set
        End Property

        Public Property ProfileMod2() As String
            Get
                Return oProfileInfo.ProfileMod2
            End Get
            Set(ByVal Value As String)
                oProfileInfo.ProfileMod2 = Value
                colProfiles(oProfileInfo.ID) = oProfileInfo
            End Set
        End Property

        Public Property User() As String
            Get
                Return oProfileInfo.User
            End Get

            Set(ByVal value As String)
                oProfileInfo.User = value
                colProfiles(oProfileInfo.ID) = oProfileInfo
            End Set
        End Property

        Public Property ProfileValue() As String
            Get
                Return oProfileInfo.ProfileValue
            End Get

            Set(ByVal value As String)
                oProfileInfo.ProfileValue = value
                colProfiles(oProfileInfo.ID) = oProfileInfo
            End Set
        End Property
        'Public ReadOnly Property EntityType() As Integer
        '    Get
        '        Return nEntityTypeID
        '    End Get
        'End Property

        Public Property Deleted() As Boolean
            Get
                Return oProfileInfo.Deleted
            End Get

            Set(ByVal value As Boolean)
                oProfileInfo.Deleted = value
                colProfiles(oProfileInfo.ID) = oProfileInfo
            End Set
        End Property

        Public Property colIsDirty() As Boolean
            Get
                Dim xProfInf As Muster.Info.ProfileInfo
                For Each xProfInf In colProfiles.Values
                    If xProfInf.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
            Set(ByVal Value As Boolean)

            End Set
        End Property

        Public Property IsDirty() As Boolean
            Get
                Return oProfileInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oProfileInfo.IsDirty = value
                colProfiles(oProfileInfo.ID) = oProfileInfo
            End Set
        End Property

        Public Property CreatedBy() As String
            Get
                Return oProfileInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oProfileInfo.CreatedBy = Value
            End Set
        End Property

        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oProfileInfo.CreatedOn
            End Get
        End Property

        Public Property ModifiedBy() As String
            Get
                Return oProfileInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oProfileInfo.ModifiedBy = Value
            End Set
        End Property

        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oProfileInfo.ModifiedOn
            End Get
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        ' Obtains and returns an entity as called for by name
        Public Function Retrieve(ByVal FullKey As String, Optional ByVal ShowDeleted As Boolean = False) As Muster.Info.ProfileInfo
            Try
                If colProfiles.Contains(FullKey) Then
                    oProfileInfo = colProfiles.Item(FullKey)
                    If oProfileInfo.IsAgedData = True And oProfileInfo.IsDirty = False Then
                        colProfiles.Remove(oProfileInfo)
                        Return Retrieve(FullKey)
                    Else
                        Return oProfileInfo
                    End If

                Else
                    Dim strArray() As String
                    Dim colTemp As Muster.Info.ProfileCollection
                    Dim oInfTemp As Muster.Info.ProfileInfo
                    strArray = FullKey.Split("|")
                    'Try
                    colTemp = oProfileDB.DBGetByKey(strArray, ShowDeleted)
                    'Catch ex As Exception
                    '    Throw ex
                    'End Try
                    For Each oInfTemp In colTemp.Values
                        colProfiles.Add(oInfTemp)
                    Next

                    If colProfiles.Count > 0 Then
                        If colTemp.Count > 1 Then
                            strArray = colTemp.GetKeys()
                            Return Retrieve(strArray(0), ShowDeleted)
                        Else
                            If colTemp.Count = 1 Then
                                '
                                ' Have to return the info as the FullKey passed might
                                '   actually be a partial, which would result in an
                                '   infinitely recursive loop.
                                '
                                oProfileInfo = colTemp.Item(colTemp.GetKeys(0))
                                Return oProfileInfo 'colTemp.Item(colTemp.GetKeys(0))
                            Else
                                Return Nothing
                            End If
                        End If
                    Else
                        Return Nothing
                    End If
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        'Saves the data in the current Info object
        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal OverrideRights As Boolean = False)
            Try
                oProfileDB.Put(oProfileInfo, moduleID, staffID, returnVal, OverrideRights)
                If Not returnVal = String.Empty Then
                    Exit Sub
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

            oProfileInfo.Archive()
            oProfileInfo.IsDirty = False
        End Sub
#End Region
#Region "Collection Operations"
        Function GetAll() As Muster.Info.ProfileCollection
            colProfiles.Clear()
            Try
                colProfiles = oProfileDB.GetAllInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Return colProfiles
        End Function
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oProfileInf As Muster.Info.ProfileInfo)

            Try
                oProfileInfo = oProfileInf
                colProfiles.Add(oProfileInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oProfileInf As Muster.Info.ProfileInfo)

            Try
                colProfiles.Remove(oProfileInf)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

            Throw New Exception("Profile Info " & oProfileInf.ID & " is not in the collection of profile data.")

        End Sub
        Public Function Items() As Muster.Info.ProfileCollection
            Return colProfiles
        End Function
        Public Function Values() As ICollection
            Return colProfiles.Values
        End Function
        Public Function Contains(ByVal FullKey As String) As Muster.Info.ProfileInfo
            oProfileInfo.ID = FullKey
            If colProfiles.Contains(oProfileInfo) Then
                Return oProfileInfo
            End If
        End Function
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim xProfInf As MUSTER.Info.ProfileInfo
            For Each xProfInf In colProfiles.Values
                If xProfInf.IsDirty Then
                    oProfileInfo = xProfInf
                    Me.Save(moduleID, staffID, returnVal)
                End If
            Next
            'Adam Nall - Added to change the pProfile isDirty. It changes this on the save 
            '            of the single profileInfo but not the parent class isDirty
            Me.colIsDirty = False
        End Sub
        Public Function GetNext() As String
            Return GetNextPrev(1)
        End Function
        Public Function GetPrevious() As String
            Return GetNextPrev(-1)
        End Function
        Private Function GetNextPrev(ByVal direction As Integer) As String
            Dim strArr() As String = colProfiles.GetKeys()
            'Dim nArr(strArr.GetUpperBound(0)) As Integer
            'Dim y As String
            'For Each y In strArr
            '    nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            'Next
            'nArr.Sort(nArr)
            colIndex = Array.BinarySearch(strArr, Me.ID.ToString)
            If colIndex + direction > -1 And _
                colIndex + direction <= strArr.GetUpperBound(0) Then
                Return colProfiles.Item(strArr.GetValue(colIndex + direction)).ID.ToString
            Else
                Return colProfiles.Item(strArr.GetValue(colIndex)).ID.ToString
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oProfileInfo = New MUSTER.Info.ProfileInfo
            colProfiles.Clear()
        End Sub
        Public Sub Reset()
            Dim xProfInf As MUSTER.Info.ProfileInfo
            If Not colProfiles.Values Is Nothing Then
                For Each xProfInf In colProfiles.Values
                    If xProfInf.IsDirty Then
                        xProfInf.Reset()
                    End If
                Next
            Else
                oProfileInfo.Reset()
            End If
        End Sub
        Public Function GetTable(ByVal strSQL As String) As DataTable
            Try
                Return oProfileDB.DBGetDS(strSQL).Tables(0)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
#End Region
#Region "Miscellaneous Operations"
        'Returns a datatable of the entities in the collection
        Public Function ProfileTable() As DataTable

            Dim oProfileInfoLocal As Muster.Info.ProfileInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable

            Try
                tbEntityTable.Columns.Add("User ID", System.Type.GetType("System.String"))
                tbEntityTable.Columns.Add("Profile Key", System.Type.GetType("System.String"))
                tbEntityTable.Columns.Add("Profile Modifier 1", System.Type.GetType("System.String"))
                tbEntityTable.Columns.Add("Profile Modifier 2", System.Type.GetType("System.String"))
                tbEntityTable.Columns.Add("Profile Value", System.Type.GetType("System.String"))
                tbEntityTable.Columns.Add("Deleted", System.Type.GetType("System.Boolean"))
                tbEntityTable.Columns.Add("Created By", System.Type.GetType("System.String"))
                tbEntityTable.Columns.Add("Created On", System.Type.GetType("System.DateTime"))
                tbEntityTable.Columns.Add("Modified By", System.Type.GetType("System.String"))
                tbEntityTable.Columns.Add("Modified On", System.Type.GetType("System.DateTime"))

                For Each oProfileInfoLocal In colProfiles.Values
                    dr = tbEntityTable.NewRow()
                    dr("User ID") = oProfileInfoLocal.User
                    dr("Profile Key") = oProfileInfoLocal.ProfileKey
                    dr("Profile Modifier 1") = oProfileInfoLocal.ProfileMod1
                    dr("Profile Modifier 2") = oProfileInfoLocal.ProfileMod2
                    dr("Profile Value") = oProfileInfoLocal.ProfileValue
                    dr("Deleted") = oProfileInfoLocal.Deleted
                    dr("Created By") = oProfileInfoLocal.CreatedBy
                    dr("Created On") = oProfileInfoLocal.CreatedOn
                    dr("Modified By") = oProfileInfoLocal.ModifiedBy
                    dr("Modified On") = oProfileInfoLocal.ModifiedOn
                    tbEntityTable.Rows.Add(dr)
                Next
                Return tbEntityTable
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        ''P1 12/10/04
        'Added by Padmaja to get a set of specified records
        Public Function ProfileComboTable(ByVal FullKey As String, Optional ByVal ShowDeleted As Boolean = False) As DataTable
            Dim strArray() As String
            Dim strSQL As String
            Dim dsSet As DataSet
            Try
                strArray = FullKey.Split("|")
                strSQL = "SELECT * FROM tblSYS_PROFILE_INFO WHERE USER_ID = '" & strArray(0) & "' AND PROFILE_KEY = '" & strArray(1) & "' "
                If strArray.Length >= 3 Then
                    If strArray(2) <> String.Empty Then
                        strSQL += " AND PROFILE_MODIFIER_1 = '" & strArray(2) & "' "
                    End If
                End If
                If strArray.Length >= 4 Then
                    If strArray(3) <> String.Empty Then
                        strSQL += " AND PROFILE_MODIFIER_2 = '" & strArray(3) & "' "
                    End If
                End If
                strSQL += IIf(Not ShowDeleted, " AND DELETED <> 1", "")
                dsSet = oProfileDB.DBGetDS(strSQL)
                Return dsSet.Tables(0)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function ProfileKeyValuesTable(ByVal FullKey As String, Optional ByVal ShowDeleted As Boolean = False) As DataTable
            Dim strArray() As String
            Dim strSQL As String
            Dim dsSet As DataSet
            Try
                strArray = FullKey.Split("|")
                strSQL = "SELECT * FROM tblSYS_PROFILE_INFO WHERE USER_ID = '" & strArray(0) & "' AND PROFILE_KEY = '" & strArray(1) & "' "
                If strArray.Length >= 3 Then
                    If strArray(2) <> String.Empty Then
                        strSQL += " AND PROFILE_MODIFIER_1 = '" & strArray(2) & "' "
                    End If
                End If
                If strArray.Length >= 4 Then
                    If strArray(3) <> String.Empty Then
                        strSQL += " AND PROFILE_MODIFIER_2 = '" & strArray(3) & "' "
                    End If
                End If
                strSQL += IIf(Not ShowDeleted, " AND DELETED <> 1", "")
                dsSet = oProfileDB.DBGetDS(strSQL)
                Return dsSet.Tables(0)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function ProfileUserTable(Optional ByVal ShowDeleted As Boolean = False) As DataTable
            Dim strArray() As String
            Dim strSQL As String
            Dim dsSet As DataSet
            Try
                strSQL = "Select '' as [USER_ID] UNION Select distinct [USER_ID] from tblSYS_PROFILE_INFO"
                strSQL += IIf(Not ShowDeleted, " WHERE DELETED <> 1", "")
                dsSet = oProfileDB.DBGetDS(strSQL)
                Return dsSet.Tables(0)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function ProfileKeyTable(Optional ByVal ShowDeleted As Boolean = False) As DataTable
            Dim strArray() As String
            Dim strSQL As String
            Dim dsSet As DataSet
            Try
                strSQL = "Select '' as [PROFILE_KEY] UNION Select distinct PROFILE_KEY from tblSYS_PROFILE_INFO"
                strSQL += IIf(Not ShowDeleted, " WHERE DELETED <> 1", "")
                dsSet = oProfileDB.DBGetDS(strSQL)
                Return dsSet.Tables(0)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
#End Region
#End Region
#Region "External Event Handlers"
        Private Sub ThisProfile(ByVal bolValue As Boolean) Handles oProfileInfo.InfoBecameDirty
            '
            ' Alert the client that the current profile data has changed
            '
            RaiseEvent InfoBecameDirty(Me.IsDirty Or Me.colIsDirty)
            'RaiseEvent test()
        End Sub
        Private Sub ThisProfileCol() Handles colProfiles.InfoChanged
            '
            ' Alert the client that the current profile data has changed
            '
            RaiseEvent InfoBecameDirty(True)
            'RaiseEvent test()
        End Sub
#End Region

    End Class
End Namespace
