'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.pUserGroupMemberships
'   Provides the collection and info object to the client for manipulating
'     Assigned Memberships of User Groups.  Class is primarily intended as a helper
'     class for the pUser class. 
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'   1.0         JC      12/31/04    Original class definition.
'   1.1         AN      01/04/05    Added Try catch and Exception Handling/Logging
'   1.2         JC      01/12/05    Added function IsAMemberOf
'   1.3         MNR     01/25/05    Added GetNext(), GetPrevious() and GetNextPrev(..) functions
'   1.4         MR      03/01/05    Altered ListGroups() Function to Change the Query.

'
' Operations
' Function                      Description
' New()                         Initializes the ProfileCollection and ProfileInfo objects.
' Retrieve(Key, [ShowDeleted])  Sets the internal ProfileInfo to the ProfileInfo matching the 
'                                supplied key.  In the event that a partial key is supplied,
'                                all matching ProfileInfos are populated to the internal 
'                                ProfileCollection and the internal ProfileInfo object is
'                                set to the first member of the collection.  In either case,
'                                the internal ProfileInfo object is returned to the client.
' Add(ProfileInfo)              Adds the supplied ProfileInfo object to the internal ProfileCollection
'                                and sets the internal ProfileInfo object to same.
' Remove(ProfileInfo)           Removes the supplied ProfileInfo object from the internal ProfileCollection
'                                if it is contained by the collection.
' Items()                       Returns the internal ProfileCollection (a Dictionary object) to the client.
' Values()                      Returns the collection of ProfileInfo objects to the client (used in for..next).
' Clear()                       Sets the internal ProfileInfo object to an empty ProfileInfo and clears
'                                the internal ProfileCollection.
' Reset()                       Reverts the internal ProfileInfo object to it's state when last retrieved
'                                from or marshalled to the repository.
' Save()                        Marshalls the internal ProfileInfo to the repository.
' Flush()                       Marshalls all modified/new ProfileInfo objects in the ProfileCollection
'                                to the repository.
'
' Properties
' Name                      Description
'  UserGroup                    Gets or Sets the UserGroup of the Screens to select
'  ScreenName                   Gets the Name of the current screen
'  ScreenPermission             Gets or Sets the Permission of current screen
'  ColIsDirty                   Returns a BOOLEAN if the Collections of Params is Dirty
'  CreatedBy                    ReadOnly returns name of user that created the item
'  CreatedOn                    ReadOnly returns the date the item was created
'  ModifiedBy                   ReadOnly returns the name of the last person to modify the item
'  ModifiedOn                   ReadOnly returns the date the item was last modified
'
' Events
' Name                      Description
' ScreensChanged                Raised to indicate that state of the screen collection changed.
'                                   Returns colMemberships.colIsDirty
'-------------------------------------------------------------------------------
'
' TODO - Add to app 1/3/2005 - JVC 2
' TODO - Check properties and operations lists.
'

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pUserGroupMemberships
#Region "Private Member Variables"
        Private WithEvents colMemberships As MUSTER.Info.ProfileCollection
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private WithEvents oProfileInfo As MUSTER.Info.ProfileInfo
        Private oProfileDB As New MUSTER.DataAccess.ProfileDB
        Private strParent As String
        Private bolShowDeleted As Boolean = False
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Public Event Handlers"
        Public Event MembershipsChanged(ByVal IsDirtyState As Boolean)
#End Region
#Region "Constructors"
        Public Sub New()
            oProfileInfo = New MUSTER.Info.ProfileInfo
            colMemberships = New MUSTER.Info.ProfileCollection
        End Sub
        Public Sub New(ByVal strUserName As String)
            oProfileInfo = New MUSTER.Info.ProfileInfo
            colMemberships = New Muster.Info.ProfileCollection
            Try
                Me.Retrieve(strUserName & "|USER GROUPS")
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property UserName() As String
            Get
                Return oProfileInfo.User
            End Get
            Set(ByVal UserName As String)
                Try
                    Me.Retrieve(UserName & "|USER GROUPS")
                Catch Ex As Exception
                    If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                    Throw Ex
                End Try
            End Set
        End Property
        Public WriteOnly Property AddMembership() As String
            Set(ByVal GroupName As String)
                Try
                    Me.Retrieve(oProfileInfo.User & "|USER GROUPS|" & GroupName & "|NONE")
                    Me.oProfileInfo.ProfileValue = GroupName
                Catch Ex As Exception
                    If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                    Throw Ex
                End Try
                'Me.Items(oProfileInfo.User & "|USER GROUPS|" & GroupName & "|NONE").ProfileValue = GroupName
            End Set
        End Property
        Public WriteOnly Property RemoveMembership() As String
            Set(ByVal GroupName As String)
                Try
                    Me.Retrieve(oProfileInfo.User & "|USER GROUPS|" & GroupName & "|NONE")
                    Me.oProfileInfo.ProfileValue = "NONE"
                Catch Ex As Exception
                    If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                    Throw Ex
                End Try
                'Me.Items(oProfileInfo.User & "|USER GROUPS|" & GroupName & "|NONE").ProfileValue = "NONE"
            End Set
        End Property
        Public WriteOnly Property SetCurrent() As String
            Set(ByVal GroupName As String)
                Try
                    Me.oProfileInfo = Me.Items(oProfileInfo.User & "|USER GROUPS|" & GroupName & "|NONE")
                Catch Ex As Exception
                    If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                    Throw Ex
                End Try
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oProfileInfo.IsDirty
            End Get
            Set(ByVal Value As Boolean)
                oProfileInfo.IsDirty = Value
            End Set
        End Property
        Public Property colIsDirty() As Boolean
            Get
                Dim xProfInf As MUSTER.Info.ProfileInfo
                For Each xProfInf In colMemberships.Values
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

        Public Property ShowDeleted() As Boolean
            Get
                Return bolShowDeleted
            End Get
            Set(ByVal Value As Boolean)
                bolShowDeleted = Value
            End Set
        End Property
        Public ReadOnly Property CreatedBy() As String
            Get
                Return oProfileInfo.CreatedBy
            End Get
        End Property

        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oProfileInfo.CreatedOn
            End Get
        End Property

        Public ReadOnly Property ModifiedBy() As String
            Get
                Return oProfileInfo.ModifiedBy
            End Get
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
        Public Sub Populate(ByVal UserName As String)
            Try
                Me.Retrieve(UserName & "|USER GROUPS")
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub

        Public Function Retrieve(ByVal FullKey As String, Optional ByVal ShowDeleted As Boolean = False) As Muster.Info.ProfileInfo
            Dim strUserName As String = String.Empty
            Dim arrGroups As New ArrayList
            Dim arrNewGroups As New ArrayList
            Dim strGroupsaMemberOf As String = String.Empty
            oProfileInfo = Nothing
            Try
                If colMemberships.Contains(FullKey) Then
                    oProfileInfo = colMemberships.Item(FullKey)
                    Return oProfileInfo
                Else
                    Dim strArray() As String
                    Dim colTemp As Muster.Info.ProfileCollection
                    Dim oInfTemp As Muster.Info.ProfileInfo
                    strArray = FullKey.Split("|")
                    colTemp = oProfileDB.DBGetByKey(strArray, ShowDeleted)
                    For Each oInfTemp In colTemp.Values
                        If Not colMemberships.Contains(oInfTemp.ID) Then
                            '
                            ' Add the Group to the list of groups accounted for
                            '
                            If arrGroups.IndexOf(oInfTemp.ProfileMod1) = -1 Then
                                arrGroups.Add(oInfTemp.ProfileMod1)
                            End If
                            colMemberships.Add(oInfTemp)
                            strGroupsaMemberOf += "'" + oInfTemp.ProfileMod1 + "', "
                            oProfileInfo = oInfTemp
                        End If
                    Next

                    If strGroupsaMemberOf.Length > 0 Then
                        strGroupsaMemberOf = " WHERE GROUP_NAME NOT IN (" + strGroupsaMemberOf.Substring(0, strGroupsaMemberOf.Length - 2) + ") AND "
                    Else
                        strGroupsaMemberOf = " WHERE "
                    End If

                    If strArray.Length > 0 Then
                        strUserName = strArray(0)
                    End If
                    '
                    ' Now build the list of groups which the user is NOT a member of...
                    '
                    Dim strSQLString As String = "SELECT DISTINCT GROUP_NAME FROM tblSYS_USER_GROUPS " + strGroupsaMemberOf + " DELETED = 0"
                    Dim oDT As New DataTable
                    Dim oRow As DataRow
                    Dim nIndex As Int16
                    oDT = oProfileDB.DBGetDS(strSQLString).Tables(0)
                    For Each oRow In oDT.Rows
                        If arrGroups.Count = 0 Then
                            arrNewGroups.Add(oRow.Item("GROUP_NAME"))
                        Else
                            If arrGroups.IndexOf(oRow.Item("GROUP_NAME")) = -1 Then
                                arrNewGroups.Add(oRow.Item("GROUP_NAME"))
                            End If
                        End If
                    Next
                    arrGroups = Nothing
                    For nIndex = 0 To arrNewGroups.Count - 1
                        Dim oInfTemp2 As New Muster.Info.ProfileInfo(strUserName, "USER GROUPS", arrNewGroups(nIndex), "NONE", "NONE", False, "SYSTEM", Now, "SYSTEM", Now)
                        oInfTemp2.IsDirty = False
                        colMemberships.Add(oInfTemp2)
                    Next
                    oProfileInfo = colMemberships.Item(colMemberships.GetKeys(0))
                    Return oProfileInfo
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function

        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try
                oProfileDB.Put(oProfileInfo, moduleID, staffID, returnVal)
                If Not returnVal = String.Empty Then
                    Exit Sub
                End If
                oProfileInfo.Archive()
                oProfileInfo.IsDirty = False
                RaiseEvent MembershipsChanged(Me.colIsDirty)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Function IsAMemberOf(ByVal strValue As String) As Boolean
            If Not colMemberships.Contains(strValue) Then
                Return False
            Else
                Return colMemberships(strValue).ProfileValue = colMemberships(strValue).ProfileMod1
            End If
        End Function
#End Region
#Region "Collection Operations"
        Public Function DumpStatus() As String
            Dim s As String = String.Empty
            Dim x As MUSTER.Info.ProfileInfo
            For Each x In colMemberships.Values
                s += x.User + "|" + x.ProfileKey + "|" + x.ProfileMod1 + "|" + IIf(x.IsDirty, "Yes", "No") + vbCrLf
            Next
            Return s
        End Function
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oProfileInf As MUSTER.Info.ProfileInfo)

            Try
                oProfileInfo = oProfileInf
                colMemberships.Add(oProfileInf)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        'Removes the entity supplied from the collection
        Private Sub Remove(ByVal oProfileInf As MUSTER.Info.ProfileInfo)

            Try
                colMemberships.Remove(oProfileInf)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

            Throw New Exception("Profile Info " & oProfileInf.ID & " is not in the collection of profile data.")

        End Sub
        Private Function Items() As Muster.Info.ProfileCollection
            Return colMemberships
        End Function
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String)
            Dim xProfInf As MUSTER.Info.ProfileInfo
            For Each xProfInf In colMemberships.Values
                If xProfInf.IsDirty Then
                    oProfileInfo = xProfInf
                    If oProfileInfo.User = String.Empty Then
                        oProfileInfo.CreatedBy = UserID
                    Else
                        oProfileInfo.ModifiedBy = UserID
                    End If
                    Me.Save(moduleID, staffID, returnVal)
                End If
            Next
            RaiseEvent MembershipsChanged(Me.colIsDirty)
        End Sub
        Public Function GetNext() As String
            Return GetNextPrev(1)
        End Function
        Public Function GetPrevious() As String
            Return GetNextPrev(-1)
        End Function
        Private Function GetNextPrev(ByVal direction As Integer) As String
            Dim strArr() As String = colMemberships.GetKeys()
            'Dim nArr(strArr.GetUpperBound(0)) As Integer
            'Dim y As String
            'For Each y In strArr
            '    nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            'Next
            'nArr.Sort(nArr)
            colIndex = Array.BinarySearch(strArr, Me.UserName.ToString)
            If colIndex + direction > -1 And _
                colIndex + direction <= strArr.GetUpperBound(0) Then
                Return colMemberships.Item(strArr.GetValue(colIndex + direction)).ID.ToString
            Else
                Return colMemberships.Item(strArr.GetValue(colIndex)).ID.ToString
            End If
        End Function
#End Region
#Region "General Operations"
        Public Function Clear()
            colMemberships.Clear()
        End Function
        Public Function Reset()
            oProfileInfo.Reset()
        End Function
        Public Function ResetCollection()
            '
            ' What is going on here???
            ' VB will not allow an in-place update of the collection members
            '   while iterating the collection.  So, build a collection of the
            '   dirty members.  Then, iterate the dirty member collection, 
            '   updating the source collection members.
            '
            ' It is an obtuse way of handling it, but it works!
            '
            Dim xInfo As Muster.Info.ProfileInfo
            Dim arrInfo As New Collection
            For Each xInfo In colMemberships.Values
                If xInfo.IsDirty Then
                    arrInfo.Add(xInfo)
                End If
            Next
            For Each xInfo In arrInfo
                xInfo.Reset()
                colMemberships.Add(xInfo)
            Next
            RaiseEvent MembershipsChanged(Me.colIsDirty)
        End Function
#End Region
#Region "Miscellaneous Operations"

        Public Function ListGroups() As DataTable
            Dim MyDT As DataTable
            Dim MyDR As DataRow
            Dim strSQL As String = "SELECT DISTINCT GROUP_NAME FROM tblSYS_USER_GROUPS WHERE DELETED <>1 ORDER BY GROUP_NAME"
            Dim dtInfo As DataSet
            Dim drInfo As DataRow
            Try
                MyDT = New DataTable
                MyDT.Columns.Add(New System.Data.DataColumn("USER_GROUP", System.Type.GetType("System.String")))
                MyDT.Columns.Add(New System.Data.DataColumn("INACTIVE", System.Type.GetType("System.Boolean")))

                dtInfo = oProfileDB.DBGetDS(strSQL)

                For Each drInfo In dtInfo.Tables(0).Rows
                    MyDR = MyDT.NewRow
                    MyDR.Item("USER_GROUP") = drInfo.Item("GROUP_NAME")
                    MyDT.Rows.Add(MyDR)
                Next
                Return MyDT
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function ListMemberships(ByVal UserName As String) As DataTable
            Dim MyDT As DataTable
            Dim MyDR As DataRow
            Dim oTempMembership As Muster.Info.ProfileInfo
            Try
                MyDT = New DataTable
                MyDT.Columns.Add(New System.Data.DataColumn("USER_GROUP", System.Type.GetType("System.String")))
                MyDT.Columns.Add(New System.Data.DataColumn("INACTIVE", System.Type.GetType("System.Boolean")))

                For Each oTempMembership In colMemberships.Values
                    If oTempMembership.User = UserName Then
                        If oTempMembership.ProfileMod1 = oTempMembership.ProfileValue Then
                            MyDR = MyDT.NewRow
                            MyDR.Item("USER_GROUP") = oTempMembership.ProfileMod1
                            MyDR.Item("INACTIVE") = oTempMembership.Deleted
                            MyDT.Rows.Add(MyDR)
                        End If
                    End If
                Next
                Return MyDT
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function ListNonMemberships(ByVal UserName As String) As DataTable
            Dim MyDT As DataTable
            Dim MyDR As DataRow
            Dim oTempMembership As Muster.Info.ProfileInfo
            Try
                MyDT = New DataTable
                MyDT.Columns.Add(New System.Data.DataColumn("USER_GROUP", System.Type.GetType("System.String")))
                MyDT.Columns.Add(New System.Data.DataColumn("INACTIVE", System.Type.GetType("System.Boolean")))

                For Each oTempMembership In colMemberships.Values
                    If oTempMembership.User = UserName Then
                        If oTempMembership.ProfileValue = "NONE" Then
                            MyDR = MyDT.NewRow
                            MyDR.Item("USER_GROUP") = oTempMembership.ProfileMod1
                            MyDR.Item("INACTIVE") = oTempMembership.Deleted
                            MyDT.Rows.Add(MyDR)
                        End If
                    End If
                Next
                Return MyDT
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function ListMembers(ByVal UserGroup As String) As DataTable
            Dim MyDT As DataTable
            Dim MyDR As DataRow
            Dim oTempMembership As Muster.Info.ProfileInfo
            Try
                MyDT = New DataTable
                MyDT.Columns.Add(New System.Data.DataColumn("USER_NAME", System.Type.GetType("System.String")))
                MyDT.Columns.Add(New System.Data.DataColumn("INACTIVE", System.Type.GetType("System.Boolean")))

                For Each oTempMembership In colMemberships.Values
                    If oTempMembership.ProfileMod1 = UserGroup Then
                        MyDR = MyDT.NewRow
                        MyDR.Item("USER_NAME") = oTempMembership.User
                        MyDR.Item("INACTIVE") = oTempMembership.Deleted
                        MyDT.Rows.Add(MyDR)
                    End If
                Next
                Return MyDT
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
#End Region
#End Region
#Region "External Event Handlers"
        Private Sub colMemberships_InfoChanged() Handles colMemberships.InfoChanged
            RaiseEvent MembershipsChanged(Me.colIsDirty)
        End Sub
        Private Sub MembershipChanged(ByVal bolState As Boolean) Handles oProfileInfo.InfoBecameDirty
            RaiseEvent MembershipsChanged(Me.colIsDirty)
        End Sub
#End Region
    End Class
End Namespace
