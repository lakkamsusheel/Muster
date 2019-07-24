'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.PropertyType
'   Provides the operations required to manipulate an Entity object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'   1.0         AN      01/21/05    Original class definition
'   1.1         MNR     01/25/05    Added GetNext(), GetPrevious() and GetNextPrev(..) functions
'   1.2         JC      02/21/05    Modified BuildPropertiesAndChildren to pass the property type ID to
'                                      BuildChildProperties and then call oPropertyDB.DBGetAllChildPropertiesByParentID
'                                       with the property type as well.
'   1.3         AB      02/22/05    Added DataAge check to the Retrieve function
'   1.4         AN      02/24/05    Removed ADD to collection on RetreiveProperty
'   1.5         JC      06/02/05    Modified getAvailableProperty to use expose new MaintenanceMode attribute.
'                                   Also changed call to DB operation to override the parent property ID (which
'                                    is normally the contained info ID) with the supplied value if the operation
'                                    is called with an explicit Parent ID value.
'
' Function          Description
' Retrieve(ID)                  Returns an Info Object requested by the int arg ID
' Retrieve(Property Name)       Returns an Info Object requested by the str arg Name
' Save()                        Saves the current Info Object
' Flush()                       Marshalls all modified/new Property Info objects in the
'                               Property Collection to the repository
' GetAll()                      Returns an PropertyTypeCollection with all Property Type objects
' Add(ID)                       Adds the Property identified by arg ID to the 
'                                   internal PropertyTypeCollection
' Add(Name)                     Adds the Property Type identified by arg NAME to the internal 
'                                   PropertyTypeCollection
' Add(Entity)                   Adds the Property Type passed as the argument to the internal 
'                                   PropertyTypeCollection
' Remove(ID)                    Removes the Property Type identified by arg ID from the internal 
'                                   PropertyTypeCollection
' Remove(NAME)                  Removes the Property Type identified by arg NAME from the 
'                                   internal PropertyTypeCollection
' PropertiesTable()             Returns a datatable containing all properties associated to the current
'                                   Property Type.
' DeletePropertyRelation(ParentID,ChildID) 
'                               Removes the realtionship record for the specified parent and child id.
' getAvaliableProperties(optional ParentID,optional ChildID)
'                               Returns a datatable containing all avaliable properties
'
' Properties
' Name                      Description
'  Entity_ID                     Gets or Sets the ID of the Entity for current Property Type
'  Property_Name                 Gets or Sets the Name of the Property Type
'  Property_ID                   Gets or Sets the ID of the Property Type
'  IsDirty                       Gets or Sets weather the current Property Type is Dirty
'  ColIsDirty                    Gets or Sets weather any Property Type in the collection is dirty
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pPropertyType
#Region "Public Events"
        Public Event PropertyTypeErr(ByVal MsgStr As String)
        Public Event PropertyTypeChanged(ByVal bolValue As Boolean)
        Public Event ColChanged(ByVal bolValue As Boolean)
        Public Event PropertyInfoChanged(ByVal bolValue As Boolean)
        Public Event PropertyColChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        Private WithEvents oPropertyTypeInfo As Muster.Info.PropertyTypeInfo
        'Private ocolPropertyTypeProperties As New Muster.Info.PropertyCollection

        Private WithEvents colPropertyTypes As Muster.Info.PropertyTypeCollection
        'Private WithEvents oProperty As New MUSTER.BusinessLogic.pProperty
        Private WithEvents oPropertyInfo As Muster.Info.PropertyInfo
        Private oPropertyTypeDB As New Muster.DataAccess.PropertyTypeDB
        Private oPropertyDB As New Muster.DataAccess.PropertymasterDB
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private nID As Int64 = -1
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
#End Region
#Region "Constructors"
        Public Sub New()
            oPropertyTypeInfo = New Muster.Info.PropertyTypeInfo
            colPropertyTypes = New Muster.Info.PropertyTypeCollection
            oPropertyInfo = New Muster.info.PropertyInfo
        End Sub
        '********************************************************
        '
        ' Overloaded NEW which will populate with a single instance
        '   of the named Template object.
        '
        '********************************************************
        Public Sub New(ByVal PropType As String)
            oPropertyTypeInfo = New Muster.Info.PropertyTypeInfo
            colPropertyTypes = New Muster.Info.PropertyTypeCollection
            oPropertyInfo = New Muster.info.PropertyInfo
            'ocolPropertyTypeProperties = New Muster.Info.PropertyCollection
            'oProperty = New MUSTER.BusinessLogic.pProperty
            Me.Retrieve(PropType)
            'TODO Add method that populates all the properties for this type
            'We will need to then take these properties and put them in the local col
            'A method that clears all the items out of the local collections
            'Then does a get for the new type is nes.
            'This method will loop through a dataset of relations and retrieve all thsese
            'from the collection and since the retreive mthod will add anything not in the
            'collection we dont have to worry about it.

            'Me.colPropertyTypes.GetAllfor RETURNS A COLECTION.
            'Me.ocolPropertyTypeProperties.Clear()
            'Me.ocolPropertyTypeProperties = Me.oProperty.GetAllbyPropType(PropType)
        End Sub
#End Region
#Region "Exposed Attributed - PropertyInfo"
        Public Property PropertyID() As String
            Get
                Return oPropertyInfo.ID
            End Get
            Set(ByVal Value As String)
                oPropertyInfo.ID = Value
            End Set
        End Property
        Public Property PropertyPropType_ID() As Long
            Get
                Return oPropertyInfo.PropType_ID
            End Get
            Set(ByVal Value As Long)
                oPropertyInfo.PropType_ID = Value
            End Set
        End Property
        Public Property PropertyName() As String
            Get
                Return oPropertyInfo.Name
            End Get
            Set(ByVal Value As String)
                oPropertyInfo.Name = Value
            End Set
        End Property
        Public Property PropertyPropDesc() As String
            Get
                Return oPropertyInfo.PropDesc
            End Get
            Set(ByVal Value As String)
                oPropertyInfo.PropDesc = Value
            End Set
        End Property
        Public Property PropertyPropPos() As Integer
            Get
                Return oPropertyInfo.PropPos
            End Get
            Set(ByVal Value As Integer)
                oPropertyInfo.PropPos = Value
            End Set
        End Property
        Public Property PropertyBUSINESSTAG() As Integer
            Get
                Return oPropertyInfo.BUSINESSTAG
            End Get
            Set(ByVal Value As Integer)
                oPropertyInfo.BUSINESSTAG = Value
            End Set
        End Property
        Public Property PropertyPropIsActive() As Boolean
            Get
                Return oPropertyInfo.PropIsActive
            End Get
            Set(ByVal Value As Boolean)
                oPropertyInfo.PropIsActive = Value
            End Set
        End Property
        Public Property PropertyProp_Type() As String
            Get
                Return oPropertyInfo.Prop_Type
            End Get
            Set(ByVal Value As String)
                oPropertyInfo.Prop_Type = Value
            End Set
        End Property
        Public Property CreatedBy() As String
            Get
                Return oPropertyInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oPropertyInfo.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oPropertyInfo.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oPropertyInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oPropertyInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oPropertyInfo.ModifiedOn
            End Get
        End Property

#End Region
#Region "Exposed Attributes"
        Public Property Properties() As Muster.Info.PropertyCollection
            Get
                Return Me.oPropertyTypeInfo.Properties 'oProperty
            End Get
            Set(ByVal Value As Muster.Info.PropertyCollection)
                Me.oPropertyTypeInfo.Properties = Value
            End Set
        End Property

        Public Property EntityId() As Integer
            Get
                Return oPropertyTypeInfo.EntityId
            End Get
            Set(ByVal Value As Integer)
                oPropertyTypeInfo.EntityId = Value
            End Set
        End Property
        ' The "common name" used to refer to the property type.
        Public Property Name() As String
            Get
                Return oPropertyTypeInfo.Name
            End Get
            Set(ByVal Value As String)
                oPropertyTypeInfo.Name = Value
            End Set
        End Property
        ' The identifier used by the repository to identify the class of the property type
        Public Property ID() As Long
            Get
                Return oPropertyTypeInfo.ID
            End Get
            Set(ByVal Value As Long)
                oPropertyTypeInfo.ID = Value
            End Set
        End Property

        Public Property Created_By() As String
            Get
                Return oPropertyTypeInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oPropertyTypeInfo.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property Created_On() As Date
            Get
                Return oPropertyTypeInfo.CreatedOn
            End Get
        End Property
        Public Property Modified_By() As String
            Get
                Return oPropertyTypeInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oPropertyTypeInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property Modified_On() As Date
            Get
                Return oPropertyTypeInfo.ModifiedOn
            End Get
        End Property

        Public Property IsDirty() As Boolean
            Get
                Return oPropertyTypeInfo.IsDirty Or Me.PropertiescolIsDirty
            End Get

            Set(ByVal value As Boolean)
                oPropertyTypeInfo.IsDirty = Boolean.Parse(value)
            End Set
        End Property
        Public Property colIsDirty() As Boolean
            Get
                Dim xPropertyTypeInfo As MUSTER.Info.PropertyTypeInfo
                For Each xPropertyTypeInfo In colPropertyTypes.Values
                    If xPropertyTypeInfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
            Set(ByVal Value As Boolean)
                oPropertyTypeInfo.IsDirty = Value
            End Set
        End Property

        Public Property PropertiescolIsDirty() As Boolean
            Get
                Dim xPropertyInfo As MUSTER.Info.PropertyInfo
                For Each xPropertyInfo In Me.Properties.Values
                    If xPropertyInfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
            Set(ByVal Value As Boolean)
                oPropertyInfo.IsDirty = Value
            End Set
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        'Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal ID As Integer) As Muster.Info.PropertyTypeInfo
            Dim oPropertyTypeInfoLocal As MUSTER.Info.PropertyTypeInfo
            Dim bolDataAged As Boolean = False
            Try
                For Each oPropertyTypeInfoLocal In colPropertyTypes.Values
                    If oPropertyTypeInfoLocal.ID = ID Then
                        If oPropertyTypeInfoLocal.IsAgedData = True And oPropertyTypeInfoLocal.IsDirty = False Then
                            bolDataAged = True
                            Exit For
                        Else
                            oPropertyTypeInfo = oPropertyTypeInfoLocal
                            'Me.ocolPropertyTypeProperties.Clear()
                            'Me.ocolPropertyTypeProperties = Me.oProperty.GetAllbyPropType(ID)
                            'Me.oProperty.Clear()                        
                            ' Me.oProperty.GetAllbyPropType(ID)

                            BuildPropertiesAndChildren(ID)
                            Return oPropertyTypeInfo
                        End If
                    End If
                Next

                If bolDataAged Then
                    colPropertyTypes.Remove(oPropertyTypeInfoLocal)
                End If

                oPropertyTypeInfo = oPropertyTypeDB.DBGetPropertyTypeByID(ID)
                If oPropertyTypeInfo.ID = 0 Then
                    oPropertyTypeInfo.ID = nID
                    nID -= 1
                End If
                colPropertyTypes.Add(oPropertyTypeInfo)
                'Me.ocolPropertyTypeProperties.Clear()
                'Me.ocolPropertyTypeProperties = Me.oProperty.GetAllbyPropType(ID)
                'Me.oProperty.Clear()
                'Me.oProperty.GetAllbyPropType(ID)
                BuildPropertiesAndChildren(ID)
                Return oPropertyTypeInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function Retrieve(ByVal TemplateName As String) As MUSTER.Info.PropertyTypeInfo
            Dim bolDataAged As Boolean = False
            Try
                oPropertyTypeInfo = Nothing
                If colPropertyTypes.Contains(TemplateName) Then
                    oPropertyTypeInfo = colPropertyTypes(TemplateName)
                    If oPropertyTypeInfo.IsAgedData = True And oPropertyTypeInfo.IsDirty = False Then
                        bolDataAged = True
                    End If
                End If

                If bolDataAged Then
                    colPropertyTypes.Remove(oPropertyTypeInfo)
                End If

                If Not colPropertyTypes.Contains(TemplateName) Then
                    If oPropertyTypeInfo Is Nothing Then
                        oPropertyTypeInfo = New MUSTER.Info.PropertyTypeInfo
                    End If
                    oPropertyTypeInfo = oPropertyTypeDB.DBGetPropertyTypeByName(TemplateName)
                    If Not colPropertyTypes.Contains(oPropertyTypeInfo) Then
                        colPropertyTypes.Add(oPropertyTypeInfo)
                    End If
                End If

                'Me.ocolPropertyTypeProperties.Clear()
                'Me.ocolPropertyTypeProperties = Me.oProperty.GetAllbyPropType(TemplateName)
                'Me.oProperty.Clear()
                'Me.oProperty.GetAllbyPropType(TemplateName)
                BuildPropertiesAndChildren(TemplateName)
                Return oPropertyTypeInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function

        Public Function RetrieveChildProperties(ByVal PropertyID As Integer) As MUSTER.Info.PropertyCollection
            Try
                Dim LocalPropertyInfo As MUSTER.Info.PropertyInfo
                For Each LocalPropertyInfo In Me.Properties.Values
                    If LocalPropertyInfo.ID = PropertyID Then
                        Return LocalPropertyInfo.ChildProperties
                    End If
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function

        Public Function RetrieveProperty(ByVal PropertyID As Integer) As MUSTER.Info.PropertyInfo
            Try
                Dim LocalPropertyInfo As MUSTER.Info.PropertyInfo
                For Each LocalPropertyInfo In Me.Properties.Values
                    If LocalPropertyInfo.ID = PropertyID Then
                        oPropertyInfo = LocalPropertyInfo
                        Return LocalPropertyInfo
                    End If
                Next

                LocalPropertyInfo = oPropertyDB.DBGetPropertyByID(PropertyID)
                'Me.Properties.Add(LocalPropertyInfo)
                Return LocalPropertyInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function

        Public Function RetrieveChildProperty(ByVal PropertyID As Integer, ByVal oPropertyinfo As MUSTER.Info.PropertyInfo) As MUSTER.Info.PropertyInfo
            Try
                Dim LocalPropertyInfo As MUSTER.Info.PropertyInfo
                For Each LocalPropertyInfo In oPropertyinfo.ChildProperties.Values
                    If LocalPropertyInfo.ID = PropertyID Then
                        oPropertyinfo = LocalPropertyInfo
                        Return LocalPropertyInfo
                    End If
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function

        Private Function BuildPropertiesAndChildren(ByVal ID As Integer)
            'Me.Properties = oPropertyDB.DBGetAllPropByTypeID(ID)
            'BuildChildProperties(Me.Properties, Me.ID)

            BuildPropertyList(oPropertyDB.DBGetAllPropByTypeID_DS(ID))
        End Function

        Private Function BuildPropertiesAndChildren(ByVal TypeName As String)
            Me.Properties = oPropertyDB.DBGetAllPropByType(TypeName)
            BuildChildProperties(Me.Properties, Me.ID)
        End Function

        Private Function BuildChildProperties(ByRef Properties As MUSTER.Info.PropertyCollection, ByVal ID As Int64)
            Dim LocalPropertyInfo As MUSTER.Info.PropertyInfo
            For Each LocalPropertyInfo In Properties.Values
                LocalPropertyInfo.ChildProperties = oPropertyDB.DBGetAllChildPropertiesByParentID(LocalPropertyInfo.ID, ID)
            Next
        End Function

        Private Function BuildPropertyList(ByRef DS As DataSet)
            If DS.Tables(0).Rows.Count > 0 Then
                Dim row As DataRow
                Me.Properties.Clear()
                For Each row In DS.Tables(0).Rows
                    Dim LocalPropertyInfo As MUSTER.Info.PropertyInfo

                    If row.Item("PARENT_ID") = 0 Then
                        'property is a PARENT
                        'Me.RetrieveProperty(row.Item("PROPERTY_ID"))
                        LocalPropertyInfo = New MUSTER.Info.PropertyInfo(row)
                        Me.Properties.Add(LocalPropertyInfo)
                    ElseIf row.Item("PARENT_ID") > 0 Then
                        'property is a CHILD
                        LocalPropertyInfo = Me.RetrieveProperty(row.Item("PARENT_ID"))
                        LocalPropertyInfo.ChildProperties.Add(New MUSTER.Info.PropertyInfo(row))
                    End If
                Next
            End If
        End Function

        'Saves the data in the current Info object
        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String)
            Dim strModuleName As String = String.Empty
            Try
                'If Me.ValidateData(strModuleName) Then
                If oPropertyTypeInfo.ID <= 0 Then
                    oPropertyTypeInfo.CreatedBy = UserID
                Else
                    oPropertyTypeInfo.ModifiedBy = UserID
                End If
                oPropertyTypeDB.Put(oPropertyTypeInfo, moduleID, staffID, returnVal)
                If Not returnVal = String.Empty Then
                    Exit Sub
                End If
                oPropertyTypeInfo.Archive()
                oPropertyTypeInfo.IsDirty = False
                Dim oProperty As MUSTER.Info.PropertyInfo
                For Each oProperty In oPropertyTypeInfo.Properties.Values
                    If oProperty.IsDirty Then
                        If oPropertyTypeInfo.ID <= 0 Then
                            oProperty.CreatedBy = UserID
                        Else
                            oProperty.ModifiedBy = UserID
                        End If
                        Me.oPropertyDB.Put(oProperty, moduleID, staffID, returnVal)
                        If Not returnVal = String.Empty Then
                            Exit Sub
                        End If
                        oProperty.Archive()
                        oProperty.IsDirty = False
                    End If
                Next
                RaiseEvent PropertyTypeChanged(oPropertyTypeInfo.IsDirty)
                'Me.oProperty.Flush()
                ''TODO Add Collection save
                'End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Validates the data before saving
        Public Function ValidateData(Optional ByVal [module] As String = "Registration") As Boolean
            Dim errStr As String = ""
            Dim validateSuccess As Boolean = False
            '********************************************************
            '
            ' Sample validation code below.  Modify as necessary to
            '  handle rules for the object.  Note that errStr should
            '  be built with all failed validation reasons as it is
            '  raised to the consumer for display to the user.  The
            '  same goes for the boolean validateSuccess.  Since
            '  validateSuccess ASSUMES failure, it must be set to
            '  TRUE if all validations are passed successfully.
            '********************************************************
            Try
                Select Case [module]
                    Case "Registration"

                        ' if any validations failed
                        Exit Select
                    Case "Technical"

                        ' if any validations failed
                        Exit Select
                End Select
                If errStr.Length > 0 Or Not validateSuccess Then
                    RaiseEvent PropertyTypeErr(errStr)
                End If
                Return validateSuccess
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region
#Region "Collection Operations"
        'Gets all the info
        Function GetAll() As Muster.Info.PropertyTypeCollection
            Try
                colPropertyTypes.Clear()
                colPropertyTypes = oPropertyTypeDB.DBGetAllPropertyType
                Return colPropertyTypes
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal ID As Integer)
            Try
                oPropertyTypeInfo = oPropertyTypeDB.DBGetPropertyTypeByID(ID)
                If oPropertyTypeInfo.ID = 0 Then
                    oPropertyTypeInfo.ID = nID
                    nID -= 1
                End If
                colPropertyTypes.Add(oPropertyTypeInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oTemplate As Muster.Info.PropertyTypeInfo)
            Try
                oPropertyTypeInfo = oTemplate
                colPropertyTypes.Add(oPropertyTypeInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Integer)
            Dim myIndex As Int16 = 1
            Dim oPropertyTypeInfoLocal As Muster.Info.PropertyTypeInfo

            Try
                For Each oPropertyTypeInfoLocal In colPropertyTypes.Values
                    If oPropertyTypeInfoLocal.ID = ID Then
                        colPropertyTypes.Remove(oPropertyTypeInfoLocal)
                        Exit Sub
                    End If
                    myIndex += 1
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("Pipe " & ID.ToString & " is not in the collection of Pipes.")
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oTemplate As Muster.Info.PropertyTypeInfo)
            Try
                colPropertyTypes.Remove(oTemplate)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("Template " & oTemplate.ID & " is not in the collection of Templates.")
        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String)
            Dim xPropertyTypeInfo As MUSTER.Info.PropertyTypeInfo
            For Each xPropertyTypeInfo In colPropertyTypes.Values
                Me.oPropertyTypeInfo = xPropertyTypeInfo
                If Me.IsDirty Then 'xPropertyTypeInfo.IsDirty Then
                    oPropertyTypeInfo = xPropertyTypeInfo
                    '********************************************************
                    '
                    ' Note that if there are contained objects and the respective
                    '  contained object collections are dirty, then the contained
                    '  collections MUST BE FLUSHED before this object can be
                    '  saved.  Otherwise, there is a risk that an attempt will
                    '  be made to insert a new object to the repository without
                    '  corresponding contained information being present which 
                    '  may, in turn, cause a foreign key violation!
                    '
                    '********************************************************
                    'Me.oProperty.Flush()
                    Me.Save(moduleID, staffID, returnVal, UserID)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
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
            Dim strArr() As String = colPropertyTypes.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colPropertyTypes.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            Else
                Return colPropertyTypes.Item(nArr.GetValue(colIndex)).ID.ToString
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oPropertyTypeInfo = New MUSTER.Info.PropertyTypeInfo
        End Sub
        Public Sub Reset()
            Dim xPropertyTypeInfo As MUSTER.Info.PropertyTypeInfo
            For Each xPropertyTypeInfo In colPropertyTypes.Values
                Me.oPropertyTypeInfo = xPropertyTypeInfo
                oPropertyTypeInfo.Reset()
                Dim oProperty As MUSTER.Info.PropertyInfo
                For Each oProperty In oPropertyTypeInfo.Properties.Values
                    If oProperty.IsDirty Then
                        oProperty.Reset()
                    End If
                Next
            Next
        End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function PropertiesTable(Optional ByVal PropType As Int64 = 0) As DataTable
            Dim dr As DataRow

            Dim tblPropTypes As New DataTable
            Dim LocPropertyInfo As Muster.Info.PropertyInfo
            Dim colSubProps As Collection
            Dim nIndex As String = IIf(PropType = 0, "PRIMARY", PropType.ToString & "S")

            Try
                tblPropTypes.Columns.Add("Property ID")
                tblPropTypes.Columns.Add("Property Name")
                tblPropTypes.Columns.Add("Parent Property")
                tblPropTypes.Columns.Add("Property Position")
                tblPropTypes.Columns.Add("PropType_ID")
                tblPropTypes.Columns.Add("Property Active", GetType(System.Boolean))
                tblPropTypes.Columns.Add("Created By")
                tblPropTypes.Columns.Add("Created On")
                tblPropTypes.Columns.Add("Modified By")
                tblPropTypes.Columns.Add("Modified On")
                tblPropTypes.Columns("Property ID").DefaultValue = 0
                tblPropTypes.Columns("Parent Property").DefaultValue = 0
                tblPropTypes.Columns("PropType_ID").DefaultValue = 0
                tblPropTypes.Columns("Property Active").DefaultValue = True
                tblPropTypes.Columns("Created By").DefaultValue = ""
                tblPropTypes.Columns("Created On").DefaultValue = ""
                tblPropTypes.Columns("Modified By").DefaultValue = ""
                tblPropTypes.Columns("Modified On").DefaultValue = ""


                'tblPropTypes.Constraints.Add("MustNotOverlap", tblPropTypes.Columns("Property Position"), False)
                'tblPropTypes.Columns("Property Name").AllowDBNull = False
                'tblPropTypes.Columns("Property Position").AllowDBNull = False
                'tblPropTypes.Columns("Property Position").Unique = True




                If Not Me.Properties Is Nothing Then
                    If Me.Properties.Count > 0 Then
                        Try
                            For Each LocPropertyInfo In Me.Properties.Values
                                dr = tblPropTypes.NewRow()
                                dr("Property ID") = LocPropertyInfo.ID
                                dr("Property Name") = LocPropertyInfo.Name
                                'dr("Parent Property") = IIf(oProperty. = 0, Nothing, oProperty.PropParent)
                                dr("Property Position") = LocPropertyInfo.PropPos
                                dr("Property Active") = LocPropertyInfo.PropIsActive
                                dr("PropType_ID") = LocPropertyInfo.PropType_ID
                                dr("Created By") = LocPropertyInfo.CreatedBy
                                'dr("Created On") = IIf(Utils.IsDateNull(oProperty.CreatedOn), Nothing, oProperty.CreatedOn)
                                dr("Modified By") = LocPropertyInfo.ModifiedBy
                                'dr("Modified On") = IIf(Utils.IsDateNull(oProperty.ModifiedOn), Nothing, oProperty.ModifiedOn)

                                tblPropTypes.Rows.Add(dr)
                            Next
                        Catch ex As Exception
                            '
                            ' There's nothing to do - if an exception was generated, then
                            '  simply return an empty datatable
                            '
                            Throw ex
                        End Try
                    End If
                End If
                Return tblPropTypes
            Catch ex As Exception
                Throw ex
            End Try

        End Function
        Public Function PutPropertyRelation(ByVal dtPropertyRel As DataTable, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String) As Boolean
            Dim success As Boolean = False
            success = Me.oPropertyTypeDB.PutPropertyRelation(dtPropertyRel, moduleID, staffID, returnVal, UserID)
            If Not returnVal = String.Empty Then
                Return False
                Exit Function
            End If

            Return success

        End Function
        Public Function DeletePropertyRelation(ByVal Parent_Property_ID As Int64, ByVal Child_Property_ID As Int64) As Boolean
            'This function will delete the property relation record for the child_id specified and current property type
            'Once the property relation record is deleted the item should be removed from the collection.
            Me.RetrieveProperty(Parent_Property_ID).ChildProperties.Remove(Me.RetrieveChildProperty(Child_Property_ID, Me.RetrieveProperty(Parent_Property_ID)))
            Return Me.oPropertyTypeDB.DeletePropertyRelation(Parent_Property_ID, Child_Property_ID)
            'TODO ADD CODE TO REMOVE ITEM FROM COLLECTION
            'Me.ocolPropertyTypeProperties.Remove()
        End Function
        Public Function getAvailableProperties(Optional ByVal Parent_Property_ID As Int64 = 0, Optional ByVal Child_Parent_ID As Int64 = 0, Optional ByVal MaintenanceMode As Boolean = False) As DataTable
            '
            ' Modified 6/2/05 - JVC - Aded MaintenanceMode parameter to call
            '                         Also added IIF to override internal property ID with supplied ID if > 0
            '
            Return Me.oPropertyTypeDB.getAvailableProperties(IIf(Parent_Property_ID > 0, Parent_Property_ID, Me.nID), Child_Parent_ID, MaintenanceMode)
        End Function
        Public Function GetByEntity(ByVal EntityID As Int64) As DataTable
            Dim dtPropertyTypeTable As New DataTable
            Dim drPTNRow As DataRow
            Dim dr As DataRow
            Dim dtPropertyTypeList As DataTable
            dtPropertyTypeTable.Columns.Add("Property Type ID")
            dtPropertyTypeTable.Columns.Add("Property Type Name")
            Try
                dtPropertyTypeList = Me.oPropertyTypeDB.getPropertyTypesbyEntity(EntityID)
                For Each dr In dtPropertyTypeList.Rows
                    Dim oPropType As MUSTER.Info.PropertyTypeInfo
                    If dr.Item("PROPERTY_TYPE_ID").ToString <> "" Then
                        'oPropType = Me.Retrieve(CInt(dr.Item("PROPERTY_TYPE_ID").ToString))
                        drPTNRow = dtPropertyTypeTable.NewRow
                        drPTNRow.Item("Property Type ID") = dr.Item("PROPERTY_TYPE_ID").ToString 'oPropType.ID
                        drPTNRow.Item("Property Type Name") = dr.Item("PROPERTY_TYPE_NAME").ToString 'oPropType.Name
                        dtPropertyTypeTable.Rows.Add(drPTNRow)
                    End If
                Next
                Return (dtPropertyTypeTable)
            Catch ex As Exception
                Throw ex
            End Try
        End Function
        'Public Sub GetProperty(ByVal ID As Integer)
        '    Me.oPropertyInfo = Me.Properties.Retrieve(ID)
        'End Sub
        Public Function GetDS(ByVal strSQL As String) As DataSet
            Return oPropertyTypeDB.DBGetDS(strSQL)
        End Function
#End Region
#End Region
#Region "External Event Handlers"
        Private Sub PropertyTypeInfoChanged(ByVal bolValue As Boolean) Handles oPropertyTypeInfo.PropertyTypeChanged
            RaiseEvent PropertyTypeChanged(bolValue)
        End Sub
        Private Sub PropertyTypeColChanged(ByVal bolValue As Boolean) Handles colPropertyTypes.ColChanged
            RaiseEvent ColChanged(bolValue)
        End Sub

        'Private Sub PropertyColChangedsub() Handles oProperty.ColChanged
        '    RaiseEvent PropertyColChanged(True)
        'End Sub

        Private Sub oPropertyInfoChanged(ByVal bolValue As Boolean) Handles oPropertyInfo.PropertyChanged
            RaiseEvent PropertyInfoChanged(bolValue)
        End Sub

#End Region
    End Class
End Namespace
