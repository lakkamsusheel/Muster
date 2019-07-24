'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.Property
'   Provides the operations required to manipulate property object(s).
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'   1.0         AN      01/21/05    Original class definition
'   1.1         MNR     01/25/05    Added GetNext(), GetPrevious() and GetNextPrev(..) functions
'   1.2         AB      02/22/05    Added DataAge check to the Retrieve function
'
' Function                      Description

' Retrieve(ID)                  Returns an Info Object requested by the int arg ID
' Retrieve(Property Name)       Returns an Info Object requested by the str arg Name
' Save()                        Saves the current Info Object
' Flush()                       Marshalls all modified/new Property Info objects in the
'                               Property Collection to the repository
' GetAll()                      Returns an PropertyCollection with all Property objects
' GetAllbyPropType(ID)          Returns an PropertyCollection with all Property objects for specified Prop Type ID
' GetAllbyPropType(Name)        Returns an PropertyCollection with all Property objects for specified Prop Type Name
' Add(ID)                       Adds the Property identified by arg ID to the 
'                                   internal PropertiesCollection
' Add(Name)                     Adds the Property identified by arg NAME to the internal 
'                                   PropertiesCollection
' Add(Entity)                   Adds the Property passed as the argument to the internal 
'                                   PropertiesCollection
' Remove(ID)                    Removes the Property identified by arg ID from the internal 
'                                   PropertiesCollection
' Remove(NAME)                  Removes the Property identified by arg NAME from the 
'                                   internal PropertiesCollection
' Values()          Returns the collection of Properties in the PropertiesCollection
' Count()          Returns the count of Properties in the PropertiesCollection
'
'
' Properties
' Name                      Description

'  Property_ID                   Gets or Sets the ID of the Property
'  PropType_ID                   Gets or Sets the Property Type ID of the Property
'  Property_Name                 Gets or Sets the Name of the Property
'  PropDesc                      Gets or Sets the Description of the Property
'  PropPos                       Gets or sets the Position of the Property
'  BUSINESSTAG                   Gets or Sets the Business Tag
'  PropIsActive                  Gets or Sets weather the Property is ACTIVE
'  Prop_Type                     Gets or Sets the Property Type
'  IsDirty                       Gets or Sets weather the current Property is Dirty
'  ColIsDirty                    Gets or Sets weather any Property in the collection is dirty
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pProperty
#Region "Public Events"
        Public Event TemplateErr(ByVal MsgStr As String)
        Public Event PropertyChanged(ByVal bolValue As Boolean)
        Public Event ColChanged()
#End Region
#Region "Private Member Variables"
        Private WithEvents oPropertyInfo As Muster.Info.PropertyInfo
        Private WithEvents colProperties As Muster.Info.PropertyCollection
        Private oPropertyDB As New Muster.DataAccess.PropertymasterDB
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private nID As Int64 = -1
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
#End Region
#Region "Constructors"
        Public Sub New()
            oPropertyInfo = New Muster.Info.PropertyInfo
            colProperties = New Muster.Info.PropertyCollection
        End Sub
        '********************************************************
        '
        ' Overloaded NEW which will populate with a single instance
        '   of the named Template object.
        '
        '********************************************************
        Public Sub New(ByVal TemplateName As String)
            oPropertyInfo = New Muster.Info.PropertyInfo
            colProperties = New Muster.Info.PropertyCollection
            Me.Retrieve(TemplateName)
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As String
            Get
                Return oPropertyInfo.ID
            End Get
            Set(ByVal Value As String)
                oPropertyInfo.ID = Value
            End Set
        End Property
        Public Property PropType_ID() As Long
            Get
                Return oPropertyInfo.PropType_ID
            End Get
            Set(ByVal Value As Long)
                oPropertyInfo.PropType_ID = Value
            End Set
        End Property
        Public Property Name() As String
            Get
                Return oPropertyInfo.Name
            End Get
            Set(ByVal Value As String)
                oPropertyInfo.Name = Value
            End Set
        End Property
        Public Property PropDesc() As String
            Get
                Return oPropertyInfo.PropDesc
            End Get
            Set(ByVal Value As String)
                oPropertyInfo.PropDesc = Value
            End Set
        End Property
        Public Property PropPos() As Integer
            Get
                Return oPropertyInfo.PropPos
            End Get
            Set(ByVal Value As Integer)
                oPropertyInfo.PropPos = Value
            End Set
        End Property
        Public Property BUSINESSTAG() As Integer
            Get
                Return oPropertyInfo.BUSINESSTAG
            End Get
            Set(ByVal Value As Integer)
                oPropertyInfo.BUSINESSTAG = Value
            End Set
        End Property
        Public Property PropIsActive() As Boolean
            Get
                Return oPropertyInfo.PropIsActive
            End Get
            Set(ByVal Value As Boolean)
                oPropertyInfo.PropIsActive = Value
            End Set
        End Property
        Public Property Prop_Type() As String
            Get
                Return oPropertyInfo.Prop_Type
            End Get
            Set(ByVal Value As String)
                oPropertyInfo.Prop_Type = Value
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oPropertyInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oPropertyInfo.IsDirty = Boolean.Parse(value)
            End Set
        End Property
        Public Property colIsDirty() As Boolean
            Get
                Dim xPropertyInfo As Muster.Info.PropertyInfo
                For Each xPropertyInfo In colProperties.Values
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
        Public ReadOnly Property PropCollection() As Muster.Info.PropertyCollection
            Get
                Return Me.colProperties
            End Get
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        'Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal ID As Integer) As Muster.Info.PropertyInfo
            Dim oPropertyInfoLocal As MUSTER.Info.PropertyInfo
            Dim bolDataAged As Boolean = False
            Try
                For Each oPropertyInfoLocal In colProperties.Values
                    If oPropertyInfoLocal.ID = ID Then
                        If oPropertyInfoLocal.IsAgedData = True And oPropertyInfoLocal.IsDirty = False Then
                            bolDataAged = True
                            Exit For
                        Else
                            oPropertyInfo = oPropertyInfoLocal
                            Return oPropertyInfo
                        End If
                    End If
                Next
                If bolDataAged Then
                    colProperties.Remove(oPropertyInfoLocal)
                End If
                oPropertyInfo = oPropertyDB.DBGetPropertyByID(ID)
                If oPropertyInfo.ID = 0 Then
                    oPropertyInfo.ID = nID
                    nID -= 1
                End If
                colProperties.Add(oPropertyInfo)
                Return oPropertyInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function Retrieve(ByVal TemplateName As String) As MUSTER.Info.PropertyInfo
            Dim bolDataAged As Boolean = False
            Try
                oPropertyInfo = Nothing
                If colProperties.Contains(TemplateName) Then
                    oPropertyInfo = colProperties(TemplateName)
                    If oPropertyInfo.IsAgedData = True And oPropertyInfo.IsDirty = False Then
                        bolDataAged = True
                    Else
                        Return oPropertyInfo
                    End If
                End If
                ' If data is old, remove from the collection
                If bolDataAged Then
                    colProperties.Remove(oPropertyInfo)
                End If
                If oPropertyInfo Is Nothing Then
                    oPropertyInfo = New MUSTER.Info.PropertyInfo
                End If
                ' Get data from the DB
                oPropertyInfo = oPropertyDB.DBGetPropertyByName(TemplateName)
                colProperties.Add(oPropertyInfo)

                Return oPropertyInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        'Saves the data in the current Info object
        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim strModuleName As String = String.Empty
            Try

                'If Me.ValidateData(strModuleName) Then
                oPropertyDB.Put(oPropertyInfo, moduleID, staffID, returnVal)
                oPropertyInfo.Archive()
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
                    RaiseEvent TemplateErr(errStr)
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
        Function GetAll() As Muster.Info.PropertyCollection
            Try
                colProperties.Clear()
                colProperties = oPropertyDB.DBGetAllProperties
                Return colProperties
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetAllbyPropType(ByVal typeid As Long) As Muster.Info.PropertyCollection
            Try
                colProperties.Clear()
                colProperties = oPropertyDB.DBGetAllPropByTypeID(typeid)
                Me.oPropertyInfo = colProperties.Item(colProperties.GetKeys(0))
                Return colProperties
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetAllbyPropType(ByVal typename As String) As Muster.Info.PropertyCollection
            Try
                colProperties.Clear()
                colProperties = oPropertyDB.DBGetAllPropByType(typename)
                Me.oPropertyInfo = colProperties.Item(colProperties.GetKeys(0))
                Return colProperties
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal ID As Integer)
            Try
                oPropertyInfo = oPropertyDB.DBGetPropertyByID(ID)
                If oPropertyInfo.ID = 0 Then
                    oPropertyInfo.ID = nID
                    nID -= 1
                End If
                colProperties.Add(oPropertyInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oTemplate As Muster.Info.PropertyInfo)
            Try
                oPropertyInfo = oTemplate
                colProperties.Add(oPropertyInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Integer)
            Dim myIndex As Int16 = 1
            Dim oPropertyInfoLocal As Muster.Info.PropertyInfo

            Try
                For Each oPropertyInfoLocal In colProperties.Values
                    If oPropertyInfoLocal.ID = ID Then
                        colProperties.Remove(oPropertyInfoLocal)
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
        Public Sub Remove(ByVal oTemplate As Muster.Info.PropertyInfo)
            Try
                colProperties.Remove(oTemplate)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("Template " & oTemplate.ID & " is not in the collection of Templates.")
        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String)
            Dim xPropertyInfo As MUSTER.Info.PropertyInfo
            For Each xPropertyInfo In colProperties.Values
                If xPropertyInfo.IsDirty Then
                    oPropertyInfo = xPropertyInfo
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
                    If oPropertyInfo.ID <= 0 Then
                        oPropertyInfo.CreatedBy = UserID
                    Else
                        oPropertyInfo.ModifiedBy = UserID
                    End If
                    Me.Save(moduleID, staffID, returnVal)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                End If
            Next
        End Sub
        Public Function Values() As ICollection
            Return colProperties.Values
        End Function
        Public Function Count() As Integer
            Return colProperties.Count
        End Function
        Public Function GetNext() As String
            Return GetNextPrev(1)
        End Function
        Public Function GetPrevious() As String
            Return GetNextPrev(-1)
        End Function
        Private Function GetNextPrev(ByVal direction As Integer) As String
            Dim strArr() As String = colProperties.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colProperties.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            Else
                Return colProperties.Item(nArr.GetValue(colIndex)).ID.ToString
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oPropertyInfo = New MUSTER.Info.PropertyInfo
        End Sub
        Public Sub Reset()
            oPropertyInfo.Reset()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        'TODO Remove this Function
        Public Function PutProperties(ByRef dtProperties As DataTable, ByVal Property_Type_ID As Integer)
            Try
                Me.oPropertyDB.PutProperties(dtProperties, Property_Type_ID)
            Catch ex As Exception
                Throw ex
            End Try
        End Function

        Public Function GetPropertyNameByID(ByVal nVal As Int16) As String
            Try
                Dim strPropName As String
                strPropName = Me.oPropertyDB.DBGetPropertyNameByID(nVal)
                Return strPropName
            Catch ex As Exception

            End Try

        End Function
        Public Function PropertiesTable(Optional ByVal PropType As Int64 = 0) As DataTable
            Dim dr As DataRow

            Dim tblPropTypes As New DataTable
            Dim oPropertyInfo As Muster.Info.PropertyInfo
            Dim colSubProps As Collection
            Dim nIndex As String = IIf(PropType = 0, "PRIMARY", PropType.ToString & "S")

            Try
                tblPropTypes.Columns.Add("Property ID")
                tblPropTypes.Columns.Add("Property Name")
                tblPropTypes.Columns.Add("Parent Property")
                tblPropTypes.Columns.Add("Property Position", GetType(System.Int32))
                tblPropTypes.Columns.Add("Property Active", GetType(System.Boolean))
                tblPropTypes.Columns.Add("Created By")
                tblPropTypes.Columns.Add("Created On")
                tblPropTypes.Columns.Add("Modified By")
                tblPropTypes.Columns.Add("Modified On")
                tblPropTypes.Columns("Property ID").DefaultValue = 0
                tblPropTypes.Columns("Parent Property").DefaultValue = 0
                tblPropTypes.Columns("Property Active").DefaultValue = True
                tblPropTypes.Columns("Created By").DefaultValue = ""
                tblPropTypes.Columns("Created On").DefaultValue = ""
                tblPropTypes.Columns("Modified By").DefaultValue = ""
                tblPropTypes.Columns("Modified On").DefaultValue = ""


                'tblPropTypes.Constraints.Add("MustNotOverlap", tblPropTypes.Columns("Property Position"), False)
                'tblPropTypes.Columns("Property Name").AllowDBNull = False
                'tblPropTypes.Columns("Property Position").AllowDBNull = False
                tblPropTypes.Columns("Property Position").Unique = True




                If Not colProperties Is Nothing Then
                    If colProperties.Count > 0 Then
                        Try
                            'colSubProps = colProperties(nIndex)
                            For Each oPropertyInfo In colProperties.Values
                                dr = tblPropTypes.NewRow()
                                dr("Property ID") = oPropertyInfo.ID
                                dr("Property Name") = oPropertyInfo.Name
                                dr("Parent Property") = oPropertyInfo.Parent_ID
                                dr("Property Position") = oPropertyInfo.PropPos
                                dr("Property Active") = oPropertyInfo.PropIsActive
                                dr("Created By") = oPropertyInfo.CreatedBy
                                dr("Created On") = oPropertyInfo.CreatedOn
                                dr("Modified By") = oPropertyInfo.ModifiedBy
                                dr("Modified On") = oPropertyInfo.ModifiedOn
                                tblPropTypes.Rows.Add(dr)
                            Next
                        Catch ex As Exception
                            '
                            ' There's nothing to do - if an exception was generated, then
                            '  simply return an empty datatable
                            '
                        End Try
                    End If
                End If
                Return tblPropTypes
            Catch ex As Exception
                Throw ex
            End Try

        End Function
#End Region
#End Region
#Region "External Event Handlers"
        Private Sub PropertyInfoChanged(ByVal bolValue As Boolean) Handles oPropertyInfo.PropertyChanged
            RaiseEvent PropertyChanged(bolValue)
        End Sub
        Private Sub PropertyColChanged() Handles colProperties.InfoChanged
            RaiseEvent ColChanged()
        End Sub
#End Region
    End Class
End Namespace
