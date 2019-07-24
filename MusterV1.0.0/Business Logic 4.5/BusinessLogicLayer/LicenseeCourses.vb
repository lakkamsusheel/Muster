'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.LicenseeCourses
'   Provides the operations required to manipulate an Entity object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0                              Original class definition
'  1.1        MR        06/04/05    Added and Modified Functions and Attributes
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
' NOTE: This file to be used as LicenseeCourses to build other objects.
'       Replace keyword "LicenseeCourses" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pLicenseeCourses
#Region "Public Events"
        Public Event LicenseeCoursesErr(ByVal MsgStr As String)
        Public Event LicenseeCoursesChanged(ByVal bolValue As Boolean)
        Public Event ColChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("LicenseeCourses").ID
        Private WithEvents oLicenseeCoursesInfo As MUSTER.Info.LicenseeCourseInfo
        Private WithEvents colLicenseeCourses As MUSTER.Info.LicenseeCourseCollection
        Private oLicenseeCoursesDB As New MUSTER.DataAccess.LicenseeCourseDB
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private nID As Int64 = -1
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
#End Region
#Region "Constructors"
        Public Sub New()
            oLicenseeCoursesInfo = New MUSTER.Info.LicenseeCourseInfo
            colLicenseeCourses = New MUSTER.Info.LicenseeCourseCollection
        End Sub
        '********************************************************
        '
        ' Overloaded NEW which will populate with a single instance
        '   of the named LicenseeCourses object.
        '
        '********************************************************
        Public Sub New(ByVal LicenseeCoursesName As String)
            oLicenseeCoursesInfo = New MUSTER.Info.LicenseeCourseInfo
            colLicenseeCourses = New MUSTER.Info.LicenseeCourseCollection
            Me.Retrieve(LicenseeCoursesName)
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Integer
            Get
                Return oLicenseeCoursesInfo.ID
            End Get
            Set(ByVal Value As Integer)
                oLicenseeCoursesInfo.ID = Integer.Parse(Value)
            End Set
        End Property
        Public Property LicenseeID() As Integer
            Get
                Return oLicenseeCoursesInfo.LicenseeID
            End Get
            Set(ByVal Value As Integer)
                oLicenseeCoursesInfo.LicenseeID = Value
            End Set
        End Property
        Public Property ProviderID() As Integer
            Get
                Return oLicenseeCoursesInfo.ProviderID
            End Get
            Set(ByVal Value As Integer)
                oLicenseeCoursesInfo.ProviderID = Value
            End Set
        End Property
        Public Property CourseTypeID() As Integer
            Get
                Return oLicenseeCoursesInfo.CourseTypeID
            End Get
            Set(ByVal Value As Integer)
                oLicenseeCoursesInfo.CourseTypeID = Value
            End Set
        End Property
        Public Property CourseDate() As DateTime
            Get
                Return oLicenseeCoursesInfo.CourseDate
            End Get
            Set(ByVal Value As DateTime)
                oLicenseeCoursesInfo.CourseDate = Value
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oLicenseeCoursesInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oLicenseeCoursesInfo.Deleted = Boolean.Parse(Value)
            End Set
        End Property
        Public Property CreatedBy() As String
            Get
                Return oLicenseeCoursesInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oLicenseeCoursesInfo.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oLicenseeCoursesInfo.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oLicenseeCoursesInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oLicenseeCoursesInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oLicenseeCoursesInfo.ModifiedOn
            End Get
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oLicenseeCoursesInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oLicenseeCoursesInfo.IsDirty = Boolean.Parse(value)
            End Set
        End Property
        Public Property colIsDirty() As Boolean
            Get
                Dim xLicenseeCoursesinfo As MUSTER.Info.LicenseeCourseInfo
                For Each xLicenseeCoursesinfo In colLicenseeCourses.Values
                    If xLicenseeCoursesinfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
            Set(ByVal Value As Boolean)
                oLicenseeCoursesInfo.IsDirty = Value
            End Set
        End Property
        Public Property LicCourseInfo() As MUSTER.Info.LicenseeCourseInfo
            Get
                Return Me.LicCourseInfo
            End Get

            Set(ByVal value As MUSTER.Info.LicenseeCourseInfo)
                Me.LicCourseInfo = value
            End Set
        End Property
        Public Property colLicCourse() As MUSTER.Info.LicenseeCourseCollection
            Get
                Return Me.colLicenseeCourses
            End Get

            Set(ByVal value As MUSTER.Info.LicenseeCourseCollection)
                Me.colLicenseeCourses = value
            End Set
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        'Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal ID As Integer) As MUSTER.Info.LicenseeCourseInfo
            Dim oLicenseeCoursesInfoLocal As MUSTER.Info.LicenseeCourseInfo
            Try
                For Each oLicenseeCoursesInfoLocal In colLicenseeCourses.Values
                    If oLicenseeCoursesInfoLocal.ID = ID Then
                        oLicenseeCoursesInfo = oLicenseeCoursesInfoLocal
                        Return oLicenseeCoursesInfo
                    End If
                Next
                oLicenseeCoursesInfo = oLicenseeCoursesDB.DBGetByID(ID)
                If oLicenseeCoursesInfo.ID = 0 Then
                    oLicenseeCoursesInfo.ID = nID
                    nID -= 1
                End If
                colLicenseeCourses.Add(oLicenseeCoursesInfo)
                Return oLicenseeCoursesInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function Retrieve(ByVal LicenseeCoursesName As String) As MUSTER.Info.LicenseeCourseInfo
            Try
                oLicenseeCoursesInfo = Nothing
                If colLicenseeCourses.Contains(LicenseeCoursesName) Then
                    oLicenseeCoursesInfo = colLicenseeCourses(LicenseeCoursesName)
                Else
                    If oLicenseeCoursesInfo Is Nothing Then
                        oLicenseeCoursesInfo = New MUSTER.Info.LicenseeCourseInfo
                    End If
                    colLicenseeCourses.Add(oLicenseeCoursesInfo)
                End If
                Return oLicenseeCoursesInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        'Saves the data in the current Info object
        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal bolValidated As Boolean = False)

            Try
                'If Me.ValidateData() Then
                Dim OldKey As String = oLicenseeCoursesInfo.ID.ToString
                oLicenseeCoursesDB.Put(oLicenseeCoursesInfo, moduleID, staffID, returnVal)
                If Not returnVal = String.Empty Then
                    Exit Sub
                End If

                If Not bolValidated Then
                    If oLicenseeCoursesInfo.ID.ToString <> OldKey Then
                        colLicCourse.ChangeKey(OldKey, oLicenseeCoursesInfo.ID.ToString)
                    End If
                End If
                oLicenseeCoursesInfo.Archive()
                oLicenseeCoursesInfo.IsDirty = False
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
                If oLicenseeCoursesInfo.ID <> 0 Then
                    If oLicenseeCoursesInfo.CourseTypeID <> 0 Then
                        If oLicenseeCoursesInfo.ProviderID <> 0 Then
                            If Date.Compare(oLicenseeCoursesInfo.CourseDate, CDate("01/01/0001")) = 0 Then
                                errStr += "Course Date cannot be empty" + vbCrLf
                                validateSuccess = False
                            Else
                                validateSuccess = True
                            End If
                        Else
                            errStr += "Provider cannot be empty" + vbCrLf
                            validateSuccess = False
                        End If
                    Else
                        errStr += "Course Type cannot be empty" + vbCrLf
                        validateSuccess = False
                    End If
                End If

                If errStr.Length > 0 Or Not validateSuccess Then
                    RaiseEvent LicenseeCoursesErr(errStr)
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
        Function GetAll(Optional ByVal nLicenseeID As Integer = 0) As MUSTER.Info.LicenseeCourseCollection
            Try
                colLicenseeCourses.Clear()
                colLicenseeCourses = oLicenseeCoursesDB.DBGetByLicenseeID(nLicenseeID)
                Return colLicenseeCourses
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal ID As Integer)
            Try
                Dim oLicCourseInfo As MUSTER.Info.LicenseeCourseInfo
                If ID = 0 Then
                    oLicCourseInfo = New MUSTER.Info.LicenseeCourseInfo
                    oLicCourseInfo.ID = nID
                    nID -= 1
                    oLicenseeCoursesInfo = oLicCourseInfo
                Else
                    oLicenseeCoursesInfo = oLicenseeCoursesDB.DBGetByID(ID)
                End If

                colLicenseeCourses.Add(oLicenseeCoursesInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oLicenseeCourses As MUSTER.Info.LicenseeCourseInfo)
            Try
                oLicenseeCoursesInfo = oLicenseeCourses
                colLicenseeCourses.Add(oLicenseeCoursesInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Integer)
            Dim myIndex As Int16 = 1
            Dim oLicenseeCoursesInfoLocal As MUSTER.Info.LicenseeCourseInfo

            Try
                For Each oLicenseeCoursesInfoLocal In colLicenseeCourses.Values
                    If oLicenseeCoursesInfoLocal.ID = ID Then
                        colLicenseeCourses.Remove(oLicenseeCoursesInfoLocal)
                        Exit Sub
                    End If
                    myIndex += 1
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("LicenseeCourses " & ID.ToString & " is not in the collection of LicenseeCourses.")
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oLicenseeCourses As MUSTER.Info.LicenseeCourseInfo)
            Try
                colLicenseeCourses.Remove(oLicenseeCourses)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("LicenseeCourses " & oLicenseeCourses.ID & " is not in the collection of LicenseeCourses.")
        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal nLicenseeID As Integer = 0)

            Dim IDs As New Collection
            Dim index As Integer
            Dim xLicenseeCoursesInfo As MUSTER.Info.LicenseeCourseInfo

            For Each xLicenseeCoursesInfo In colLicenseeCourses.Values
                If xLicenseeCoursesInfo.IsDirty Then
                    If nLicenseeID > 0 Then
                        xLicenseeCoursesInfo.LicenseeID = nLicenseeID
                    End If
                    oLicenseeCoursesInfo = xLicenseeCoursesInfo
                    IDs.Add(oLicenseeCoursesInfo.ID)
                    Me.Save(moduleID, staffID, returnVal, True)
                End If
            Next

            If Not (IDs Is Nothing) Then
                For index = 1 To IDs.Count
                    Dim colKey As String = CType(IDs.Item(index), String)
                    xLicenseeCoursesInfo = colLicenseeCourses.Item(colKey)
                    colLicenseeCourses.ChangeKey(colKey, xLicenseeCoursesInfo.ID)
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
            Dim strArr() As String = colLicenseeCourses.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colLicenseeCourses.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            Else
                Return colLicenseeCourses.Item(nArr.GetValue(colIndex)).ID.ToString
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oLicenseeCoursesInfo = New MUSTER.Info.LicenseeCourseInfo
        End Sub
        Public Sub Reset()
            oLicenseeCoursesInfo.Reset()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function CourseTable() As DataTable
            Dim oLicenseeCoursesInfoLocal As New MUSTER.Info.LicenseeCourseInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Try

                tbEntityTable.Columns.Add("CourseID", GetType(Integer))
                tbEntityTable.Columns.Add("LicenseeID", GetType(Integer))
                tbEntityTable.Columns("LicenseeID").DefaultValue = 0
                tbEntityTable.Columns.Add("Date", GetType(Date))
                tbEntityTable.Columns("Date").DefaultValue = Today.ToShortDateString
                tbEntityTable.Columns.Add("Type")
                tbEntityTable.Columns.Add("Provider")
                tbEntityTable.Columns.Add("Deleted", GetType(Boolean))
                tbEntityTable.Columns("Deleted").DefaultValue = False
                tbEntityTable.Columns.Add("Created By")
                tbEntityTable.Columns.Add("Date Created")
                tbEntityTable.Columns.Add("Last Edited By")
                tbEntityTable.Columns.Add("Date Last Edited")


                For Each oLicenseeCoursesInfoLocal In colLicenseeCourses.Values
                    dr = tbEntityTable.NewRow()
                    dr("CourseID") = oLicenseeCoursesInfoLocal.ID
                    dr("LicenseeID") = oLicenseeCoursesInfoLocal.LicenseeID
                    dr("Date") = oLicenseeCoursesInfoLocal.CourseDate
                    dr("Type") = oLicenseeCoursesInfoLocal.CourseTypeID
                    dr("Provider") = oLicenseeCoursesInfoLocal.ProviderID
                    dr("Deleted") = oLicenseeCoursesInfoLocal.Deleted
                    dr("Created By") = oLicenseeCoursesInfoLocal.CreatedBy
                    dr("Date Created") = oLicenseeCoursesInfoLocal.CreatedOn
                    dr("Last Edited By") = oLicenseeCoursesInfoLocal.ModifiedBy
                    dr("Date Last Edited") = oLicenseeCoursesInfoLocal.ModifiedOn
                    tbEntityTable.Rows.Add(dr)
                Next
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

                dsReturn = oLicenseeCoursesDB.DBGetDS(strSQL)
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
        Private Sub PLicenseeCoursesInfoChanged(ByVal bolValue As Boolean) Handles oLicenseeCoursesInfo.LicCourseInfoChanged
            RaiseEvent LicenseeCoursesChanged(bolValue)
        End Sub
        Private Sub PLicenseeCoursesColChanged(ByVal bolValue As Boolean) Handles colLicenseeCourses.LicenseeCourseColChanged
            RaiseEvent ColChanged(bolValue)
        End Sub
#End Region
    End Class
End Namespace
