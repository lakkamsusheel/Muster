'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.pScreens
'   Provides the collection and info object to the client for manipulating
'     Assigned Screens of User Groups.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'   1.0         AN      12/22/04    Original class definition.
'   1.1         JC      12/28/04    Added events for data update notification.
'                                   Altered RETRIEVE() to add any new screens
'                                   listed under user SYSTEM that are not in
'                                   the retrieved set (auto-update of sets).
'                                   Added event firing in FLUSH().
'                                   Added RESETCOLLECTION() function.
'                                   Altered LISTSCREENS() to include the DELETED attribute
'                                   in the returned datatable.
'   1.2         AN      01/03/05    Added Try catch and Exception Handling/Logging
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
'                                   Returns colScreens.colIsDirty
'-------------------------------------------------------------------------------

'
'TODO - Update Functions and Properties lists in header.
'

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pScreens
#Region "Private Member Variables"
        Private WithEvents colScreens As Muster.Info.ProfileCollection
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private WithEvents oParamInfo As Muster.Info.ProfileInfo
        Private oProfileDB As New Muster.DataAccess.ProfileDB
        Private strParent As String
        Private bolShowDeleted As Boolean = False
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Public Event Handlers"
        Public Event ScreensChanged(ByVal IsDirtyState As Boolean)
#End Region
#Region "Constructors"
        Public Sub New()
            oParamInfo = New Muster.Info.ProfileInfo
            colScreens = New Muster.Info.ProfileCollection
        End Sub
        Public Sub New(ByVal UserGroup As String)
            oParamInfo = New Muster.Info.ProfileInfo
            colScreens = New Muster.Info.ProfileCollection
            Me.UserGroup = UserGroup
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property UserGroup() As String
            Get
                Return strParent
            End Get
            Set(ByVal Value As String)
                strParent = Value
                Me.colScreens.Clear()
                Me.Retrieve(Value & "|SCREENS", False)
            End Set
        End Property

        Public ReadOnly Property ScreenName() As String
            Get
                Return oParamInfo.ProfileMod1
            End Get
        End Property

        Public Property ScreenPermission() As String
            Get
                If Not oParamInfo Is Nothing Then
                    Return oParamInfo.ProfileValue
                Else
                    Return String.Empty
                End If
            End Get

            Set(ByVal value As String)
                oParamInfo.ProfileValue = value
                colScreens(oParamInfo.ID) = oParamInfo
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oParamInfo.IsDirty
            End Get
            Set(ByVal Value As Boolean)
                oParamInfo.IsDirty = Value
            End Set
        End Property
        Public Property colIsDirty() As Boolean
            Get
                Dim xProfInf As Muster.Info.ProfileInfo
                For Each xProfInf In colScreens.Values
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
                Return oParamInfo.CreatedBy
            End Get
        End Property

        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oParamInfo.CreatedOn
            End Get
        End Property

        Public ReadOnly Property ModifiedBy() As String
            Get
                Return oParamInfo.ModifiedBy
            End Get
        End Property

        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oParamInfo.ModifiedOn
            End Get
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        ' Obtains and returns an entity as called for by name
        Public Function Retrieve(ByVal FullKey As String, Optional ByVal ShowDeleted As Boolean = False) As Muster.Info.ProfileInfo
            Dim strUserGroup As String = String.Empty
            oParamInfo = Nothing
            Try
                If colScreens.Contains(FullKey) Then
                    oParamInfo = colScreens.Item(FullKey)
                    Return oParamInfo
                Else
                    Dim strArray() As String
                    Dim colTemp As Muster.Info.ProfileCollection
                    Dim oInfTemp As Muster.Info.ProfileInfo
                    strArray = FullKey.Split("|")
                    colTemp = oProfileDB.DBGetByKey(strArray, ShowDeleted)
                    For Each oInfTemp In colTemp.Values
                        If Not colScreens.Contains(oInfTemp.ID) Then
                            colScreens.Add(oInfTemp)
                            oParamInfo = oInfTemp
                        End If
                    Next
                    If strArray.Length > 0 Then
                        strUserGroup = strArray(0)
                    End If
                    strArray = "SYSTEM|SCREENS".Split("|")
                    colTemp = oProfileDB.DBGetByKey(strArray, ShowDeleted)
                    For Each oInfTemp In colTemp.Values
                        If Not colScreens.Contains(strUserGroup & "|SCREENS|" & oInfTemp.ProfileMod1 & "|NONE") Then
                            oInfTemp.User = strUserGroup
                            colScreens.Add(oInfTemp)
                        End If
                    Next
                    oParamInfo = colScreens.Item(colScreens.GetKeys(0))
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function

        Public Sub CurrentScreens(ByVal UserGroup As String)
            Dim oLocalSysScreens As New MUSTER.BusinessLogic.pScreens(UserGroup)
            Dim oLocalSysScreen As Muster.Info.ProfileInfo
            Dim oLocalScreen As Muster.Info.ProfileInfo
            Dim strKey As String
            Try
                For Each oLocalSysScreen In oLocalSysScreens.Values
                    If Not colScreens.Contains(oLocalSysScreen.ID) Then
                        colScreens.Add(oLocalSysScreen)
                    End If
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            '    strKey = UserGroup & "|SCREENS|" & oLocalSysScreen.ProfileMod1 & "|NONE"
            '    If Not colScreens.Contains(strKey) Then
            '        oLocalScreen = New Muster.Info.ProfileInfo(oLocalSysScreen)
            '        colScreens.Add(oLocalScreen)
            '    End If
            'Next
        End Sub

        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try
                oProfileDB.Put(oParamInfo, moduleID, staffID, returnVal)
                If Not returnVal = String.Empty Then
                    Exit Sub
                End If
                oParamInfo.Archive()
                oParamInfo.IsDirty = False
                RaiseEvent ScreensChanged(colIsDirty)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
#End Region
#Region "Collection Operations"
        Function GetAll() As Muster.Info.ProfileCollection
            Try
                colScreens.Clear()
                colScreens = oProfileDB.GetAllInfo
                Return colScreens
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oProfileInf As Muster.Info.ProfileInfo)

            Try
                oParamInfo = oProfileInf
                colScreens.Add(oProfileInf)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oProfileInf As Muster.Info.ProfileInfo)

            Try
                colScreens.Remove(oProfileInf)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

            Throw New Exception("Profile Info " & oProfileInf.ID & " is not in the collection of profile data.")

        End Sub
        Public Function Items() As Muster.Info.ProfileCollection
            Return colScreens
        End Function
        Public Function Values() As ICollection
            Return colScreens.Values
        End Function
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String)
            Dim xProfInf As MUSTER.Info.ProfileInfo
            For Each xProfInf In colScreens.Values
                If xProfInf.IsDirty Then
                    oParamInfo = xProfInf
                    If oParamInfo.User = String.Empty Then
                        oParamInfo.CreatedBy = UserID
                    Else
                        oParamInfo.ModifiedBy = UserID
                    End If
                    Me.Save(moduleID, staffID, returnVal)
                End If
            Next
            RaiseEvent ScreensChanged(colIsDirty)
        End Sub
#End Region
#Region "General Operations"
        Public Sub Clear()
            oParamInfo = New MUSTER.Info.ProfileInfo
            colScreens.Clear()
        End Sub
        Public Sub Reset()
            oParamInfo.Reset()
        End Sub
        Public Sub ResetCollection()
            Dim xInfo As MUSTER.Info.ProfileInfo
            Dim arrInfo As New Collection
            For Each xInfo In colScreens.Values
                If xInfo.IsDirty Then
                    arrInfo.Add(xInfo)
                End If
            Next
            For Each xInfo In arrInfo
                xInfo.Reset()
                colScreens.Add(xInfo)
            Next
            RaiseEvent ScreensChanged(colIsDirty)
        End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function ListScreens(ByVal UserGroup As String) As DataTable
            Dim MyDT As DataTable
            Dim MyDR As DataRow
            Dim oTempScreen As Muster.Info.ProfileInfo
            Try
                MyDT = New DataTable
                MyDT.Columns.Add(New System.Data.DataColumn("SCREEN_NAME", System.Type.GetType("System.String")))
                MyDT.Columns.Add(New System.Data.DataColumn("ACCESS_MODE", System.Type.GetType("System.String")))
                MyDT.Columns.Add(New System.Data.DataColumn("INACTIVE", System.Type.GetType("System.Boolean")))

                For Each oTempScreen In colScreens.Values
                    If oTempScreen.User = UserGroup Then
                        MyDR = MyDT.NewRow
                        MyDR.Item("SCREEN_NAME") = oTempScreen.ProfileMod1
                        MyDR.Item("ACCESS_MODE") = oTempScreen.ProfileValue
                        MyDR.Item("INACTIVE") = oTempScreen.Deleted
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
        Private Sub colScreens_InfoChanged() Handles colScreens.InfoChanged
            RaiseEvent ScreensChanged(Me.colIsDirty)
        End Sub
        Private Sub ScreenChanged(ByVal bolState As Boolean) Handles oParamInfo.InfoBecameDirty
            RaiseEvent ScreensChanged(Me.colIsDirty)
        End Sub
#End Region
    End Class
End Namespace
