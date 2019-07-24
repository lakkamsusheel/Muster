'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.Profile
'   Provides the collection and info object to the client for manipulating
'     FavSearchChildInfo & FavSearchParentInfo data.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'   1.0         MR      12/5/04     Original class definition.
'   1.1         AN      01/03/05    Added Try catch and Exception Handling/Logging
'   1.2         MR      01/7/05     Added Events for data update notification.
'                                   Added firing of event in ITEM()
'                                   Added few functions to implement the functionality
'   1.3         JVC2    01/21/05    Changed RetrieveParent to Retrive
'                                   Made RetrieveChild a private operation.
'                                   Changed CPublic to IsPublic
'                                   Changed ChildInfo.ParentID to a private attribute
'                                   (the ADD method for child will set the ParentID)
'                                   Removed the Key attribute on child - no longer used.
'                                   Added ChildOrder to attributes on child.
'                                   Modified AddChild to check for ChildID of 0 rather than NOTHING
'                                     and set ParentID of ChildInfo to oParentInfo.ID
'                                   Modified RemoveParent to index into the ParentInfo collection
'                                     rather than iterating the entire ParentInfo collection.
'                                   Modified RemoveChild to index into the ChildInfo collection
'                                     rather than iterating the entire ChildInfo collection and
'                                     decrement the ordering of the criteria that have an order
'                                     greater than the order of the removed child.
'                                   Removed the Items and Values operations - they are not pertinent
'                                     to this object.
'                                   Changed Reset to ResetCollection and added new ResetParent and 
'                                     ResetChild functions.
'                                   Added data types to columns in tables returned by ParentTable and
'                                     ChildTable functions and added Criterion_Order to ChildTable
'                                   Eliminated GetTable function - redundant.  Client should set the
'                                     Parent and then use the ChildTable function.
'                                   Added ChildOrder attribute.
'                                   Changed GetAllChild operation to a Private operation.
'                                   Modified GetParentByID so that it attempts to access the repository
'                                     if the ParentInfo is not found in the Parent collection.
'                                   Modified GetChildByID so that it attempts to access the repository
'                                     if the ChildInfo is not found in the Child collection.
'   1.4         MNR     01/25/05    Added GetNext(), GetPrevious() and GetNextPrev(..) functions
'   1.5         JVC2    01/27/05    Added RetrieveChildByName and changed RetrieveChild to RetrieveChildByID
'   1.5         EN      02/01/05    Modified RemoveParent and Flush method... 
'   1.6         EN      02/02/05    Added new event and modified all events with source column
'   1.7         JVC2    02/02/05    Modified RemoveParent and RemoveChild to bypass PUT call if IDs are
'                                       < 0
'   1.8         EN     02/03/05    Added ValidateParent Function and added new Event.
'
'
' Operations
' Function                      Description
' New()                         Initializes the FavSearchParentCollection,FavSearchParentInfo,FavSearchChildCollection,FavSearchChildInfo and FavSearchDB objects.
' RetrieveParent(Key)           Sets the internal FavSearchParentInfo to the FavSearchParentInfo matching the 
'                                supplied key.  In the event that a partial key is supplied,
'                                all matching FavSearchParentInfos are populated to the internal 
'                                FavSearchParentCollection and the internal FavSearchParentInfo object is
'                                set to the first member of the collection.  In either case,
'                                the internal FavSearchParentInfo object is returned to the client.
'
' RetrieveChild(Key)           Sets the internal FavSearchChildInfo to the FavSearchChildInfo matching the 
'                                supplied key.  In the event that a partial key is supplied,
'                                all matching FavSearchChildInfos are populated to the internal 
'                                FavSearchChildCollection and the internal FavSearchChildInfo object is
'                                set to the first member of the collection.  In either case,
'                                the internal FavSearchChildInfo object is returned to the client.
'
' SaveParent()                  Marshalls the internal FavSearchParentInfo to the repository.
' SaveChild()                   Marshalls the internal FavSearchChildInfo to the repository.
' Add(FavSearchParentInfo)       Adds the supplied FavSearchParentInfo object to the internal FavSearchParentCollection
'                                and sets the internal FavSearchParentInfo object to same.
' Add(FavSearchChildInfo)       Adds the supplied FavSearchParentInfo object to the internal FavSearchParentCollection
'                                and sets the internal FavSearchParentInfo object to same.
' RemoveParent(ID)           Removes the supplied FavSearchParentInfo object from the internal FavSearchParentCollection
' RemoveChild(ID)            Removes the supplied FavSearchChildInfo object from the internal FavSearchParentCollection
'                                if it is contained by the collection.
' Items()                       Returns the internal FavSearchParentCollection (a Dictionary object) to the client.
' Values()                      Returns the collection of FavSearchParentInfo objects to the client (used in for..next).
' Flush()                       Marshalls all modified/new FavSearchParentInfo and FavSearchChildInfo objects in the FavSearchParentCollection and FavSearchChildCollection.
'                                to the repository.
' Clear()                       Sets the internal FavSearchParentInfo object to an empty ProfileInfo and clears
'                                the internal FavSearchParentCollection.
' Reset()                       Reverts the internal ProfileInfo object to it's state when last retrieved
'                                from or marshalled to the repository.
' GetByParentID(ID)             Retrieves Parent info based on the Input ID.
' GetByChildID(ID)              Retrieves Child info based on the Input ID.
' GetAll(strUserID)             Retrieve all Parent or Child based on the Input ID.
' GetAllChild(strUserID)        Retrieve Child infos based on the Input ID.
' GetTable(SeachID)             Returns a datatable containing the resultset from the supplied SQL string.
' ParentTable()                 Returns a datatable containing the parentInfos.
' ChildTable()                  Returns a datatable containing the ChildInfos.
'colFavSearch_UserChanged()     Raise an Parent Collection Event for Enabling/Disbaling UI controls.
'oInfoFavSearch_UserChanged()   Raise an Parent Info Event for Enabling/Disbaling UI controls.
'colCriteria_UserChanged()      Raise an Child Collection Event for Enabling/Disbaling UI controls.
'oInfoCriteria_UserChanged()    Raise an Child Info Event for Enabling/Disbaling UI controls.
'----------------------------------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pFavSearch
#Region "Private Member Variables"
        Friend WithEvents colParent As Muster.Info.FavSearchParentCollection
        Friend WithEvents colChildren As Muster.Info.FavSearchChildCollection
        Friend WithEvents oParentInfo As Muster.Info.FavSearchParentInfo
        Friend WithEvents oChildInfo As Muster.Info.FavSearchChildInfo
        Private oFavSearchDB As Muster.DataAccess.FavSearchDB
        Private nnewSearchID As String = String.Empty
        Private nnewCriteriaID As Integer = 0
        Private nParentID As Integer = -1
        Private nChildID As Integer = -1
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private MusterException As New MUSTER.Exceptions.MusterExceptions

#End Region
#Region "Constructors"
        Public Sub New()
            colParent = New Muster.Info.FavSearchParentCollection
            colChildren = New Muster.Info.FavSearchChildCollection
            oParentInfo = New Muster.Info.FavSearchParentInfo
            oChildInfo = New Muster.Info.FavSearchChildInfo
            oFavSearchDB = New Muster.DataAccess.FavSearchDB
        End Sub


#End Region
#Region "Public Events"
        Public Event FavSearchChanged(ByVal IsDirty As Boolean, ByVal StrSrc As String)
        Public Event eEnableDelete(ByVal nCount As Integer, ByVal StrSrc As String)
        Public Event eFavValidateDataErr(ByVal strMessage As String, ByVal StrSrc As String)
#End Region
#Region "Exposed Attributes"
#Region "FavSearchParentInfo Attributes"
        Public Property ID() As Integer
            Get
                Return oParentInfo.ID
            End Get

            Set(ByVal value As Integer)
                oParentInfo.ID = value
            End Set
        End Property
        Public Property Name() As String
            Get
                Return oParentInfo.Name
            End Get
            Set(ByVal Value As String)
                oParentInfo.Name = Value
                colParent(oParentInfo.ID) = oParentInfo
            End Set
        End Property
        Public Property SearchType() As String
            Get
                Return oParentInfo.SearchType
            End Get
            Set(ByVal Value As String)
                oParentInfo.SearchType = Value
                colParent(oParentInfo.ID) = oParentInfo
            End Set
        End Property
        Public Property LustStatus() As String
            Get
                Return oParentInfo.LustStatus
            End Get
            Set(ByVal Value As String)
                oParentInfo.LustStatus = Value
                colParent(oParentInfo.ID) = oParentInfo
            End Set
        End Property
        Public Property TankStatus() As String
            Get
                Return oParentInfo.TankStatus
            End Get
            Set(ByVal Value As String)
                oParentInfo.TankStatus = Value
                colParent(oParentInfo.ID) = oParentInfo
            End Set
        End Property
        Public Property User() As String
            Get
                Return oParentInfo.User
            End Get

            Set(ByVal value As String)
                oParentInfo.User = value
                colParent(oParentInfo.ID) = oParentInfo
            End Set
        End Property
        Public ReadOnly Property CriteriaCount() As Int32
            Get
                Return oParentInfo.NumCriteria
            End Get
        End Property
        Public Property IsPublic() As Boolean
            Get
                Return oParentInfo.IsPublic
            End Get

            Set(ByVal value As Boolean)
                oParentInfo.IsPublic = value
                colParent(oParentInfo.ID) = oParentInfo
            End Set
        End Property
        Public Property colIsDirty() As Boolean
            Get
                Dim xPrInfo As MUSTER.Info.FavSearchParentInfo
                For Each xPrInfo In colParent.Values
                    If xPrInfo.IsDirty Then
                        Return True
                        Exit Property
                    Else
                        Dim xChInfo As MUSTER.Info.FavSearchChildInfo
                        For Each xChInfo In colChildren.Values
                            If xChInfo.ID < 0 Then
                                If xChInfo.IsDirty Then
                                    Return True
                                    Exit Property
                                End If
                            Else
                                If xChInfo.ParentID = xPrInfo.ID Then
                                    If xChInfo.IsDirty Then
                                        Return True
                                        Exit Property
                                    End If
                                End If
                            End If
                        Next
                    End If
                Next
                Return False
            End Get
            Set(ByVal Value As Boolean)

            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oParentInfo.IsDirty Or oChildInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oParentInfo.IsDirty = value
                colParent(oParentInfo.ID) = oParentInfo
                oChildInfo.IsDirty = value
                colChildren(oChildInfo.ID) = oChildInfo
            End Set
        End Property
        Public ReadOnly Property CreatedBy() As String
            Get
                Return oParentInfo.CreatedBy
            End Get
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oParentInfo.CreatedOn
            End Get
        End Property
        Public ReadOnly Property ModifiedBy() As String
            Get
                Return oParentInfo.ModifiedBy
            End Get
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oParentInfo.ModifiedOn
            End Get
        End Property
#End Region
#Region "FavSearchChildInfo Attributes"
        Public Property ChildID() As Integer
            Get
                Return oChildInfo.ID
            End Get

            Set(ByVal value As Integer)
                oChildInfo.ID = value
            End Set
        End Property
        Public Property ChildOrder() As Integer
            Get
                Return oChildInfo.Order
            End Get
            Set(ByVal Value As Integer)
                oChildInfo.Order = Value
                Dim xChildInfo As Muster.Info.FavSearchChildInfo
                For Each xChildInfo In colChildren.Values
                    If xChildInfo.Order >= oChildInfo.Order And _
                       xChildInfo.ID <> oChildInfo.ID And _
                       xChildInfo.ParentID = oParentInfo.ID Then
                        xChildInfo.Order += 1
                    End If
                Next
            End Set
        End Property
        Public Property ChildName() As String
            Get
                Return oChildInfo.CriterionName
            End Get
            Set(ByVal Value As String)
                oChildInfo.CriterionName = Value
                colChildren(oChildInfo.ID) = oChildInfo
            End Set
        End Property
        Public Property CriterionValue() As String
            Get
                Return oChildInfo.CriterionValue
            End Get
            Set(ByVal Value As String)
                oChildInfo.CriterionValue = Value
                colChildren(oChildInfo.ID) = oChildInfo
            End Set
        End Property
        Public Property CriterionDataType() As String
            Get
                Return oChildInfo.CriterionDataType
            End Get
            Set(ByVal Value As String)
                oChildInfo.CriterionDataType = Value
                colChildren(oChildInfo.ID) = oChildInfo
            End Set
        End Property
        Private Property ParentID() As Integer
            Get
                Return oChildInfo.ParentID
            End Get

            Set(ByVal value As Integer)
                oChildInfo.ParentID = Integer.Parse(value)
                colChildren(oChildInfo.ID) = oChildInfo
            End Set
        End Property
        Public Property ChildIsDeleted() As Boolean
            Get
                Return oChildInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oChildInfo.Deleted = Value
            End Set
        End Property
#End Region
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        ' Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal FullKey As String) As Muster.Info.FavSearchParentInfo
            Try
                If colParent.Contains(FullKey) Then
                    oParentInfo = colParent.Item(FullKey)
                    Return oParentInfo
                Else
                    oParentInfo = oFavSearchDB.DBGetByParentID(CLng(FullKey))
                    If Not oParentInfo Is Nothing Then
                        colParent.Add(oParentInfo)
                    End If
                    Return oParentInfo
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function RetrieveChildByName(ByVal ChildName As String) As Muster.Info.FavSearchChildInfo
            Dim localInf As Muster.Info.FavSearchChildInfo
            For Each localInf In colChildren.Values
                If localInf.CriterionName = ChildName And localInf.ParentID = oParentInfo.ID Then
                    oChildInfo = localInf
                    Return oChildInfo
                    Exit Function
                End If
            Next
            Return Nothing
        End Function
        Private Function RetrieveChildByKey(ByVal FullKey As String) As Muster.Info.FavSearchChildInfo
            Dim LocalKey As String
            Try
                LocalKey = oParentInfo.ID & "|" & FullKey
                If colChildren.Contains(LocalKey) Then
                    oChildInfo = colChildren.Item(LocalKey)
                    Return oChildInfo
                Else
                    oChildInfo = oFavSearchDB.DBGetByChildID(CLng(FullKey))
                    If Not oChildInfo Is Nothing Then
                        colChildren.Add(oChildInfo)
                    End If
                    Return oChildInfo
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Saves the data in the current Info object
        Private Sub SaveChild()
            Dim oldChild As Muster.Info.FavSearchChildInfo

            Try
                oldChild = Nothing
                If oChildInfo.ID < 0 Then
                    oldChild = oChildInfo
                    oChildInfo.ID = 0
                    oChildInfo.ParentID = oParentInfo.ID
                End If

                oFavSearchDB.Put(oChildInfo)
                oChildInfo.Archive()
                oChildInfo.IsDirty = False
                If Not oldChild Is Nothing Then
                    colChildren.ChangeKey(oldChild.ParentID & "|" & oldChild.ID, oChildInfo.ParentID & "|" & oChildInfo.ID)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Added By Elango on Feb 3 2005 
        Public Function ValidateParent(Optional ByVal [module] As String = "Registration") As Boolean
            Try
                Dim errStr As String = ""
                Dim validateSuccess As Boolean = True
                Select Case [module]
                    Case "Registration"
                        If oParentInfo.ID <> 0 Then
                            If oParentInfo.Name = String.Empty Then
                                errStr += "Please Provide a Name for the Favorite Search" + vbCrLf
                                validateSuccess = False
                            End If
                            If oParentInfo.SearchType = String.Empty Then
                                errStr += "Please select a Search Type" + vbCrLf
                                validateSuccess = False
                            End If
                            If oParentInfo.User = String.Empty Then
                                errStr += "Search User cannot be obtained" + vbCrLf
                                validateSuccess = False
                            End If
                        End If
                        If errStr.Length > 0 Or Not validateSuccess Then
                            RaiseEvent eFavValidateDataErr(errStr, Me.ToString)
                        End If
                        Exit Select
                        'Case "Technical"
                    Case Else
                        validateSuccess = False
                End Select
                Return validateSuccess
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function



        'Saves the data in the current Info object
        Public Sub SaveParent()
            Dim oChild As Muster.Info.FavSearchChildInfo
            Dim oldParentId As Integer
            Try

                If ValidateParent() Then
                    oldParentId = oParentInfo.ID
                    oFavSearchDB.Put(oParentInfo)
                    If oldParentId <> oParentInfo.ID Then
                        colParent.ChangeKey(oldParentId, oParentInfo.ID)
                    End If
                    For Each oChild In colChildren.Values
                        If oChild.IsDirty And oChild.ParentID = Trim(oldParentId) Then
                            oChild.ParentID = oParentInfo.ID
                            '  oChild.ID = 0
                            oFavSearchDB.Put(oChild)
                            oChild.Archive()
                            oChild.IsDirty = False
                        End If
                    Next
                    oParentInfo.Archive()
                    oParentInfo.IsDirty = False
                End If

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
#End Region
#Region "Collection Operations"
        Public Function GetAll(ByVal sUserId As String) As Muster.Info.FavSearchParentCollection
            Try
                colParent.Clear()
                colParent = oFavSearchDB.GetAllParentInfo(sUserId)
                RaiseEvent eEnableDelete(colParent.Count, Me.ToString)
                GetAllChild(sUserId)
                Return colParent
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Private Function GetAllChild(ByVal sUserId As String)
            Try
                colChildren.Clear()
                colChildren = oFavSearchDB.GetAllChildInfo(sUserId)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetParentByID(ByVal ID As Integer) As Muster.Info.FavSearchParentInfo
            Try
                Dim xParentInfo As Muster.Info.FavSearchParentInfo
                oParentInfo = Nothing
                For Each xParentInfo In colParent.Values
                    If xParentInfo.ID = ID Then
                        oParentInfo = xParentInfo
                        Return oParentInfo
                    End If
                Next
                If oParentInfo Is Nothing Then
                    xParentInfo = oFavSearchDB.DBGetByParentID(ID)
                    If xParentInfo Is Nothing Then
                        Throw New Exception("Favorite Search ID " & ID.ToString & " not found!")
                        Return Nothing
                        Exit Function
                    End If
                    colParent.Add(xParentInfo)
                    Return GetParentByID(ID)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetChildByID(ByVal ID As Integer) As Muster.Info.FavSearchChildInfo
            Try
                Dim xChildInfo As Muster.Info.FavSearchChildInfo
                oChildInfo = Nothing
                For Each xChildInfo In colChildren.Values
                    If xChildInfo.ID = ID Then
                        oChildInfo = xChildInfo
                        Return oChildInfo
                    End If
                Next
                If oParentInfo Is Nothing Then
                    xChildInfo = oFavSearchDB.DBGetByChildID(ID)
                    If xChildInfo Is Nothing Then
                        Throw New Exception("Favorite Search Criterion " & ID.ToString & " not found!")
                        Return Nothing
                        Exit Function
                    End If
                    colChildren.Add(xChildInfo)
                    Return GetChildByID(ID)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        'Adds an ParentInfo to the collection as called for by ID
        Public Sub AddParent(ByRef oParentInf As Muster.Info.FavSearchParentInfo)
            Try
                oParentInfo = oParentInf
                If oParentInfo.ID = 0 Then
                    oParentInfo.ID = nParentID
                    nParentID -= 1
                End If
                colParent.Add(oParentInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Sub AddChild(ByRef oChildInf As Muster.Info.FavSearchChildInfo)
            Try
                oChildInfo = oChildInf
                If oChildInfo.ID = 0 Then
                    oChildInfo.ID = nChildID
                    nChildID -= 1
                End If
                oChildInfo.ParentID = oParentInfo.ID
                colChildren.Add(oChildInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity supplied from the collection
        Public Sub RemoveParent(ByVal ID As Int64)
            Try
                Dim ParentInfo As Muster.Info.FavSearchParentInfo
                ParentInfo = colParent.Item(ID)
                ParentInfo.Deleted = True
                If ParentInfo.ID > -1 Then
                    oFavSearchDB.Put(ParentInfo)
                End If
                colParent.Remove(ParentInfo)

                'Added By Elango on jan 31 2005 
                Dim nIndex As Long
                Dim arrKeys(1) As String
                Dim xChildInfo As MUSTER.Info.FavSearchChildInfo
                For Each xChildInfo In colChildren.Values
                    If xChildInfo.ParentID = ID Then
                        ReDim Preserve arrKeys(arrKeys.GetUpperBound(0) + 1)
                        arrKeys(arrKeys.GetUpperBound(0)) = xChildInfo.ID
                    End If
                Next
                For nIndex = 1 To arrKeys.GetUpperBound(0)
                    If Not arrKeys(nIndex) Is Nothing Then
                        Me.RemoveChild(arrKeys(nIndex).ToString)
                    End If
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity supplied from the collection
        Public Sub RemoveChild(ByVal ID As Int64)
            Dim MyOrder As Integer
            Try
                Dim xChildInfo As Muster.Info.FavSearchChildInfo
                xChildInfo = colChildren.Item(ID)
                MyOrder = xChildInfo.Order
                xChildInfo.Deleted = True
                If ID > -1 Then
                    oFavSearchDB.Put(xChildInfo)
                End If
                colChildren.Remove(xChildInfo)
                '
                ' Now modify the ordering of the remaining criteria
                '
                For Each xChildInfo In colChildren.Values
                    If xChildInfo.ParentID = oParentInfo.ID And xChildInfo.Order > MyOrder Then
                        xChildInfo.Order -= 1
                    End If
                Next

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Public Function Items() As Muster.Info.FavSearchParentCollection
        '    Return colParent
        'End Function
        'Public Function Values() As ICollection
        '    Return colParent.Values
        'End Function
        Public Sub Flush()
            'For header
            Dim nIndex As Long
            Dim arrKeys(1) As String
            Dim xParentInfo As Muster.Info.FavSearchParentInfo
            Dim xChildInfo As Muster.Info.FavSearchChildInfo
            For Each xParentInfo In colParent.Values
                If xParentInfo.IsDirty Then
                    ReDim Preserve arrKeys(arrKeys.GetUpperBound(0) + 1)
                    arrKeys(arrKeys.GetUpperBound(0)) = xParentInfo.ID
                Else
                    For Each xChildInfo In colChildren.Values
                        If xChildInfo.IsDirty And xChildInfo.ParentID = xParentInfo.ID Then
                            ReDim Preserve arrKeys(arrKeys.GetUpperBound(0) + 1)
                            arrKeys(arrKeys.GetUpperBound(0)) = xChildInfo.ParentID
                        End If
                    Next
                End If
            Next
            For nIndex = 1 To arrKeys.GetUpperBound(0)
                If Not arrKeys(nIndex) Is Nothing Then
                    oParentInfo = Me.Retrieve(arrKeys(nIndex).ToString)
                    Me.SaveParent()
                End If
            Next
        End Sub
        Public Function GetNext() As String
            Return GetNextPrev(1)
        End Function
        Public Function GetPrevious() As String
            Return GetNextPrev(-1)
        End Function
        Private Function GetNextPrev(ByVal direction As Integer) As String
            Dim strArr() As String = colParent.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colParent.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            Else
                Return colParent.Item(nArr.GetValue(colIndex)).ID.ToString
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oParentInfo = New Muster.Info.FavSearchParentInfo
            oChildInfo = New Muster.Info.FavSearchChildInfo
            'colParent.Clear()
            'colChildren.Clear()
        End Sub
        Public Sub ResetCollection()
            Dim xParentInfo As Muster.Info.FavSearchParentInfo
            If Not colParent.Values Is Nothing Then
                For Each xParentInfo In colParent.Values
                    If xParentInfo.IsDirty Then
                        xParentInfo.Reset()
                    End If
                Next
            End If

            Dim xChildInfo As Muster.Info.FavSearchChildInfo
            If Not colChildren.Values Is Nothing Then
                For Each xChildInfo In colChildren.Values
                    If xChildInfo.IsDirty Then
                        xChildInfo.Reset()
                    End If
                Next
            End If
        End Sub
        Public Sub ResetParent()
            oParentInfo.Reset()
            Dim xChildInfo As Muster.Info.FavSearchChildInfo
            If Not colChildren.Values Is Nothing Then
                For Each xChildInfo In colChildren.Values
                    If xChildInfo.IsDirty And xChildInfo.ParentID = oParentInfo.ID Then
                        xChildInfo.Reset()
                    End If
                Next
            End If
        End Sub
        Public Sub ResetChild()
            oChildInfo.Reset()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        'Returns a datatable of the addresses in the collection
        Public Function ParentTable() As DataTable

            Dim oParent As Muster.Info.FavSearchParentInfo
            Dim dr As DataRow
            Dim tbParentTable As New DataTable

            Try
                tbParentTable.Columns.Add("SEARCH_ID", System.Type.GetType("System.Int64"))
                tbParentTable.Columns.Add("SEARCH_NAME", System.Type.GetType("System.String"))
                tbParentTable.Columns.Add("SEARCH_USER", System.Type.GetType("System.String"))
                tbParentTable.Columns.Add("PUBLIC_FLAG", System.Type.GetType("System.Boolean"))
                tbParentTable.Columns.Add("SEARCH_TYPE", System.Type.GetType("System.String"))
                tbParentTable.Columns.Add("LUST_STATUS", System.Type.GetType("System.String"))
                tbParentTable.Columns.Add("TANK_STATUS", System.Type.GetType("System.String"))
                Dim ArrPK(1) As DataColumn
                ArrPK(1) = tbParentTable.Columns("SEARCH_NAME")
                tbParentTable.PrimaryKey = ArrPK

                For Each oParent In colParent.Values
                    dr = tbParentTable.NewRow()
                    dr("SEARCH_ID") = oParent.ID
                    dr("SEARCH_NAME") = oParent.Name
                    dr("SEARCH_USER") = oParent.User
                    dr("PUBLIC_FLAG") = oParent.IsPublic
                    dr("SEARCH_TYPE") = oParent.SearchType
                    dr("LUST_STATUS") = oParent.LustStatus
                    dr("TANK_STATUS") = oParent.TankStatus

                    tbParentTable.Rows.Add(dr)
                Next
                Return tbParentTable
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function ChildTable() As DataTable

            Dim oChild As Muster.Info.FavSearchChildInfo
            Dim dr As DataRow
            Dim tbChildTable As New DataTable

            Try
                tbChildTable.Columns.Add("CRITERION_ID", System.Type.GetType("System.Int64"))
                tbChildTable.Columns.Add("CRITERION_ORDER", System.Type.GetType("System.Int64"))
                tbChildTable.Columns.Add("SEARCH_ID", System.Type.GetType("System.Int64"))
                tbChildTable.Columns.Add("CRITERION_NAME", System.Type.GetType("System.String"))
                tbChildTable.Columns.Add("CRITERION_VALUE", System.Type.GetType("System.String"))
                tbChildTable.Columns.Add("CRITERION_DATA_TYPE", System.Type.GetType("System.String"))
                Dim ArrPK(1) As DataColumn
                ArrPK(1) = tbChildTable.Columns("CRITERION_ORDER")
                tbChildTable.PrimaryKey = ArrPK

                If Not oParentInfo Is Nothing Then
                    For Each oChild In colChildren.Values
                        If oChild.ParentID = oParentInfo.ID Then
                            dr = tbChildTable.NewRow()
                            dr("CRITERION_ID") = oChild.ID
                            dr("CRITERION_ORDER") = oChild.Order
                            dr("SEARCH_ID") = oChild.ParentID
                            dr("CRITERION_NAME") = oChild.CriterionName
                            dr("CRITERION_VALUE") = oChild.CriterionValue
                            dr("CRITERION_DATA_TYPE") = oChild.CriterionDataType
                            tbChildTable.Rows.Add(dr)
                        End If
                    Next
                End If
                Return tbChildTable
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
#End Region
#End Region
#Region "External Event Handlers"
        Private Sub colFavSearch_UserChanged() Handles colParent.FavSearchColChanged
            RaiseEvent FavSearchChanged(Me.colIsDirty, Me.ToString)
        End Sub
        Private Sub oInfoFavSearch_UserChanged(ByVal bolValue As Boolean) Handles oParentInfo.FavSearchInfoChanged
            RaiseEvent FavSearchChanged(bolValue, Me.ToString)
        End Sub
        Private Sub colCriteria_UserChanged() Handles colChildren.CriteriaColChanged
            RaiseEvent FavSearchChanged(Me.colIsDirty, Me.ToString)
        End Sub
        Private Sub oInfoCriteria_UserChanged(ByVal bolValue As Boolean) Handles oChildInfo.CriteriaInfoChanged
            RaiseEvent FavSearchChanged(bolValue, Me.ToString)
        End Sub
#End Region
    End Class
End Namespace

