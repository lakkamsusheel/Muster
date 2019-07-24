'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.pCourse
'   Provides the operations required to manipulate an Course object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'   1.0         MR     5/22/05    Original class definition.
'
' Function          Description
' Retrieve(ID)      Returns an Info Object requested by the int arg ID
' Save()            Saves the Info Object
' GetAll()          Returns a collection with all the relevant information
' Add(ID)           Adds an Info Object identified by the int arg ID
'                   to the Courses Collection
' Add(Entity)       Adds the Entity passed as an argument
'                   to the Courses Collection
' Remove(ID)        Removes an Info Object identified by the int arg ID
'                   from the Courses Collection
' Remove(Entity)    Removes the Entity passed as an argument
'                   from the Courses Collection
' Flush()           Marshalls all modified/new Onwer Info objects in the
'                   Course Collection to the repository
' EntityTable()     Returns a datatable containing all columns for the Entity
'                   objects in the Courses Collection
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
        Public Class pCourse
#Region "Public Events"
        Public Event CourseErr(ByVal MsgStr As String)
        Public Event evtCourseErr(ByVal MsgStr As String)
#End Region
#Region "Private Member Variables"
        Private oCourseInfo As New MUSTER.Info.CourseInfo
        Private colCourse As New MUSTER.Info.CourseCollection
        Private oCourseDB As New MUSTER.DataAccess.CourseDB
        Private nID As Int64 = -1
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private odtCourseDates As New DataTable
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("Course").ID
#End Region
#Region "Constructors"
        Public Sub New()
            oCourseInfo = New MUSTER.Info.CourseInfo
            colCourse = New MUSTER.Info.CourseCollection
            InitializeDtCourseDates()
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Integer
            Get
                Return oCourseInfo.ID
            End Get
            Set(ByVal Value As Integer)
                oCourseInfo.ID = Value
            End Set
        End Property
        Public Property Active() As Boolean
            Get
                Return oCourseInfo.Active
            End Get

            Set(ByVal value As Boolean)
                oCourseInfo.Active = value
            End Set
        End Property
        Public Property ProviderID() As Integer
            Get
                Return oCourseInfo.ProviderID
            End Get
            Set(ByVal Value As Integer)
                oCourseInfo.ProviderID = Value
            End Set
        End Property
        Public Property CourseTitle() As String
            Get
                Return oCourseInfo.CourseTitle
            End Get
            Set(ByVal Value As String)
                oCourseInfo.CourseTitle = Value
            End Set
        End Property
        Public Property CourseDates() As DataTable
            Get
                Return odtCourseDates
            End Get
            Set(ByVal Value As DataTable)
                odtCourseDates = Value
            End Set
        End Property
        'Public Property CourseDates() As String
        '    Get
        '        Return oCourseInfo.CourseDates
        '    End Get

        '    Set(ByVal value As String)
        '        oCourseInfo.CourseDates = value
        '    End Set
        'End Property
        'Public Property Location() As String
        '    Get
        '        Return oCourseInfo.Location
        '    End Get

        '    Set(ByVal value As String)
        '        oCourseInfo.Location = value
        '    End Set
        'End Property
        'Public Property CourseTypeID() As Integer
        '    Get
        '        Return oCourseInfo.CourseTypeID
        '    End Get

        '    Set(ByVal value As Integer)
        '        oCourseInfo.CourseTypeID = Integer.Parse(value)
        '    End Set
        'End Property
        Public Property ProviderName() As String
            Get
                Return oCourseInfo.ProviderName
            End Get

            Set(ByVal value As String)
                oCourseInfo.ProviderName = value
            End Set
        End Property
        'Public Property CreditHours() As String
        '    Get
        '        Return oCourseInfo.CreditHours
        '    End Get

        '    Set(ByVal value As String)
        '        oCourseInfo.CreditHours = value
        '    End Set
        'End Property

        Public Property CreatedBy() As String
            Get
                Return oCourseInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oCourseInfo.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oCourseInfo.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oCourseInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oCourseInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oCourseInfo.ModifiedOn
            End Get
        End Property

        Public Property Deleted() As Boolean
            Get
                Return oCourseInfo.Deleted
            End Get

            Set(ByVal value As Boolean)
                oCourseInfo.Deleted = value
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oCourseInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oCourseInfo.IsDirty = value
            End Set
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        'Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal ID As Integer) As MUSTER.Info.CourseInfo
            Dim oCourseInfoLocal As MUSTER.Info.CourseInfo
            Dim bolDataAged As Boolean = False

            Try
                For Each oCourseInfoLocal In colCourse.Values
                    If oCourseInfoLocal.ID = ID Then
                        If oCourseInfoLocal.IsDirty = False And oCourseInfoLocal.IsAgedData = True Then
                            bolDataAged = True
                        Else
                            oCourseInfo = oCourseInfoLocal
                            GetCourseDates(oCourseInfo.ID)
                            Return oCourseInfo
                        End If
                    End If
                Next
                If bolDataAged = True Then
                    colCourse.Remove(oCourseInfoLocal)
                End If
                oCourseInfo = oCourseDB.DBGetByID(ID)
                GetCourseDates(oCourseInfo.ID)
                If oCourseInfo.ID = 0 Then
                    oCourseInfo.ID = nID
                    nID -= 1
                End If
                colCourse.Add(oCourseInfo)
                Return oCourseInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Saves the data in the current Info object
        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)

            Try
                Dim OldKey As String = oCourseInfo.ID.ToString

                oCourseDB.Put(oCourseInfo, moduleID, staffID, returnVal)
                If oCourseInfo.ID.ToString <> OldKey Then
                    colCourse.ChangeKey(OldKey, oCourseInfo.ID.ToString)
                End If
                oCourseInfo.Archive()
                oCourseInfo.IsDirty = False

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Sub Save(ByRef oCourInfo As MUSTER.Info.CourseInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)

            Try
                Dim OldKey As String = oCourInfo.ID.ToString
                oCourseDB.Put(oCourInfo, moduleID, staffID, returnVal)

                If oCourInfo.ID.ToString <> OldKey Then
                    colCourse.ChangeKey(OldKey, oCourInfo.ID.ToString)
                End If
                oCourInfo.Archive()
                oCourInfo.IsDirty = False

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Function SaveCourseDates(ByVal UserID As String) As Boolean
            Dim drRow As DataRow
            Dim bolSuccess As Boolean

            Try
                For Each drRow In odtCourseDates.Rows
                    'If ValidateCoursedates() Then

                    bolSuccess = oCourseDB.DBPutCourseDates(Integer.Parse(drRow.Item("COURSEDATES_ID")), _
                                               Integer.Parse(drRow.Item("COURSE_ID")), _
                                               Integer.Parse(drRow.Item("COURSEDATES_NUMBER")), _
                                               drRow.Item("COURSE DATES"), _
                                               drRow.Item("LOCATION"), _
                                               Integer.Parse(drRow.Item("COURSE TYPE")), _
                                               drRow.Item("HOURS"), _
                                               drRow.Item("DELETED"), _
                                               UserID)
                    'Else
                    '    Return False

                    'End If
                Next

                GetCourseDates(oCourseInfo.ID)
                Return bolSuccess

            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        'Validate Course dates
        'Public Function ValidateCoursedates() As Boolean
        '    Dim drRow As DataRow
        '    Dim dtColumn As DataColumn
        '    Dim i As Integer = 1
        '    Try
        '        Dim errStr As String = ""
        '        Dim validateSuccess As Boolean = True
        '        If odtCourseDates.Rows.Count > 0 Then
        '            For Each drRow In odtCourseDates.Rows
        '                For Each dtColumn In odtCourseDates.Columns
        '                    If drRow(dtColumn) Is DBNull.Value Then
        '                        errStr += "Provide " + dtColumn.ColumnName.ToString + " on Row" + i.ToString + vbCrLf

        '                    End If

        '                Next
        '                i = i + 1
        '            Next
        '        End If


        '        End If
        '        If errStr.Length > 0 Or Not validateSuccess Then
        '            RaiseEvent evtCourseErr(errStr)
        '        End If
        '        Return validateSuccess
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Function
        ' Validates Data according to DDD Specifications
        Public Function ValidateData() As Boolean
            Try
                Dim errStr As String = ""
                Dim validateSuccess As Boolean = True

                'If oCourseInfo.ID <> 0 Then
                If oCourseInfo.CourseTitle <> String.Empty Then
                    If oCourseInfo.ProviderID > 0 Then
                        'If oCourseInfo.CourseTypeID > 0 Then
                        '    If oCourseInfo.Location <> String.Empty Then
                        '        validateSuccess = True
                        '    Else
                        '        errStr += "Location cannot be empty" + vbCrLf
                        '        validateSuccess = False
                        '    End If
                        'Else
                        '    errStr += "Course Type cannot be empty" + vbCrLf
                        '    validateSuccess = False
                        'End If
                    Else
                        errStr += "Provider cannot be empty" + vbCrLf
                        validateSuccess = False
                    End If
                Else
                    errStr += "Course Title cannot be empty" + vbCrLf
                    validateSuccess = False
                End If
                'End If
                If errStr.Length > 0 Or Not validateSuccess Then
                    RaiseEvent evtCourseErr(errStr)
                End If
                Return validateSuccess
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Private Sub InitializeDtCourseDates()
            odtCourseDates.Columns.Add("COURSE_ID")
            odtCourseDates.Columns.Add("COURSEDATES_ID")
            odtCourseDates.Columns.Add("COURSEDATES_NUMBER")
            odtCourseDates.Columns.Add("COURSE DATES")
            odtCourseDates.Columns.Add("LOCATION")
            odtCourseDates.Columns.Add("COURSE TYPE")
            odtCourseDates.Columns.Add("HOURS")
            odtCourseDates.Columns.Add("DELETED")
            'odtCourseDates.Columns.Add("CREATED_BY")
            'odtCourseDates.Columns.Add("DATE_CREATED")
            'odtCourseDates.Columns.Add("LAST_EDITED_BY")
            'odtCourseDates.Columns.Add("DATE_LAST_EDITED")
        End Sub
        Private Function GetCourseDates(Optional ByVal CourseID As Integer = 0, Optional ByVal showDeleted As Boolean = False) As DataTable
            Dim ds As DataSet
            Dim drNew As DataRow
            Try

                ds = oCourseDB.DBGetCourseDates(CourseID, showDeleted)
                odtCourseDates.Rows.Clear()
                If ds.Tables.Count > 0 Then
                    If ds.Tables(0).Rows.Count > 0 Then

                        For Each dr As DataRow In ds.Tables(0).Rows
                            drNew = odtCourseDates.NewRow
                            For Each col As DataColumn In ds.Tables(0).Columns
                                If (col.ColumnName <> "CREATED_BY" And col.ColumnName <> "DATE_CREATED" And col.ColumnName <> "LAST_EDITED_BY" And col.ColumnName <> "DATE_LAST_EDITED") Then
                                    drNew(col.ColumnName) = dr(col.ColumnName)
                                End If
                            Next
                            odtCourseDates.Rows.Add(drNew)
                        Next
                    End If
                End If
                Return odtCourseDates
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region
#Region "Collection Operations"
        Function GetAll() As MUSTER.Info.CourseCollection
            Try
                colCourse.Clear()
                colCourse = oCourseDB.GetAllInfo()
                Return colCourse
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal ID As Integer)

            Try
                oCourseInfo = oCourseDB.DBGetByID(ID)
                If oCourseInfo.ID = 0 Then
                    oCourseInfo.ID = nID
                    nID -= 1
                End If
                colCourse.Add(oCourseInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Function Add(ByRef oCourse As MUSTER.Info.CourseInfo) As Boolean
            Try
                oCourseInfo = oCourse
                If ValidateData() Then
                    colCourse.Add(oCourseInfo)
                    Return True
                Else
                    Return False
                End If

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Integer)
            Dim myIndex As Int16 = 1
            Dim oCourseInfoLocal As MUSTER.Info.CourseInfo

            Try
                For Each oCourseInfoLocal In colCourse.Values
                    If oCourseInfoLocal.ID = ID Then
                        colCourse.Remove(oCourseInfoLocal)
                        Exit Sub
                    End If
                    myIndex += 1
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("Course " & ID.ToString & " is not in the collection of Courses.")
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oCourse As MUSTER.Info.CourseInfo)
            Try
                colCourse.Remove(oCourse)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("Course " & oCourse.ID & " is not in the collection of Courses.")
        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim xCourseInfo As MUSTER.Info.CourseInfo
            For Each xCourseInfo In colCourse.Values
                If xCourseInfo.IsDirty Then
                    oCourseInfo = xCourseInfo
                    Me.Save(moduleID, staffID, returnVal)
                End If
            Next
        End Sub
#End Region
#Region "General Operations"
        Public Sub Clear()
            oCourseInfo = New MUSTER.Info.CourseInfo
        End Sub
        Public Sub Reset()
            oCourseInfo.Reset()
        End Sub
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
        Public Function ListCourseTitles() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vCOM_COURSETITLE")
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

                dsReturn = oCourseDB.DBGetDS(strSQL)
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
#Region "Miscellaneous Operations"

#End Region
#End Region
    End Class
End Namespace

