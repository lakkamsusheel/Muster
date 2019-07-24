'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.LicenseeCourseTest
'   Provides the operations required to manipulate an Entity object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0                              Original class definition
'  1.1        MR        06/05/05    Added and Modified Functions and Attributes
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
' NOTE: This file to be used as LicenseeCourseTest to build other objects.
'       Replace keyword "LicenseeCourseTest" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pLicenseeCourseTest
#Region "Public Events"
        Public Event LicenseeCourseTestsErr(ByVal MsgStr As String)
        Public Event LicenseeCourseTestsChanged(ByVal bolValue As Boolean)
        Public Event ColChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("LicenseeCoursesTest").ID
        Private WithEvents oLicenseeCourseTestInfo As MUSTER.Info.LicenseeCourseTestInfo
        Private WithEvents colLicenseeCoursesTest As MUSTER.Info.LicenseeCourseTestCollection
        Private oLicenseeCourseTestDB As New MUSTER.DataAccess.LicenseCoursesTestsDB
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private nID As Int64 = -1
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
#End Region
#Region "Constructors"
        Public Sub New()
            oLicenseeCourseTestInfo = New MUSTER.Info.LicenseeCourseTestInfo
            colLicenseeCoursesTest = New MUSTER.Info.LicenseeCourseTestCollection
        End Sub
        '********************************************************
        '
        ' Overloaded NEW which will populate with a single instance
        '   of the named LicenseeCourseTest object.
        '
        '********************************************************
        Public Sub New(ByVal LicenseeCourseTestName As String)
            oLicenseeCourseTestInfo = New MUSTER.Info.LicenseeCourseTestInfo
            colLicenseeCoursesTest = New MUSTER.Info.LicenseeCourseTestCollection
            Me.Retrieve(LicenseeCourseTestName)
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Integer
            Get
                Return oLicenseeCourseTestInfo.ID
            End Get
            Set(ByVal Value As Integer)
                oLicenseeCourseTestInfo.ID = Integer.Parse(Value)
            End Set
        End Property
        Public Property LicenseeID() As Integer
            Get
                Return oLicenseeCourseTestInfo.LicenseeID
            End Get
            Set(ByVal Value As Integer)
                oLicenseeCourseTestInfo.LicenseeID = Value
            End Set
        End Property
        Public Property CourseTypeID() As Integer
            Get
                Return oLicenseeCourseTestInfo.CourseTypeID
            End Get
            Set(ByVal Value As Integer)
                oLicenseeCourseTestInfo.CourseTypeID = Value
            End Set
        End Property
        Public Property TestDate() As DateTime
            Get
                Return oLicenseeCourseTestInfo.TestDate
            End Get
            Set(ByVal Value As DateTime)
                oLicenseeCourseTestInfo.TestDate = Value
            End Set
        End Property
        Public Property StartTime() As String
            Get
                Return oLicenseeCourseTestInfo.StartTime
            End Get
            Set(ByVal Value As String)
                oLicenseeCourseTestInfo.StartTime = Value
            End Set
        End Property
        Public Property TestScore() As Integer
            Get
                Return oLicenseeCourseTestInfo.TestScore
            End Get
            Set(ByVal Value As Integer)
                oLicenseeCourseTestInfo.TestScore = Value
            End Set
        End Property
        Public Property CreatedBy() As String
            Get
                Return oLicenseeCourseTestInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oLicenseeCourseTestInfo.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oLicenseeCourseTestInfo.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oLicenseeCourseTestInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oLicenseeCourseTestInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oLicenseeCourseTestInfo.ModifiedOn
            End Get
        End Property

        Public Property Deleted() As Boolean
            Get
                Return oLicenseeCourseTestInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oLicenseeCourseTestInfo.Deleted = Boolean.Parse(Value)
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oLicenseeCourseTestInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oLicenseeCourseTestInfo.IsDirty = Boolean.Parse(value)
            End Set
        End Property
        Public Property colIsDirty() As Boolean
            Get
                Dim xLicenseeCourseTestInfo As MUSTER.Info.LicenseeCourseTestInfo
                For Each xLicenseeCourseTestInfo In colLicenseeCoursesTest.Values
                    If xLicenseeCourseTestInfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
            Set(ByVal Value As Boolean)
                oLicenseeCourseTestInfo.IsDirty = Value
            End Set
        End Property
        Public Property LicCourseTestInfo() As MUSTER.Info.LicenseeCourseTestInfo
            Get
                Return Me.LicCourseTestInfo
            End Get

            Set(ByVal value As MUSTER.Info.LicenseeCourseTestInfo)
                Me.LicCourseTestInfo = value
            End Set
        End Property
        Public Property colLicCourseTest() As MUSTER.Info.LicenseeCourseTestCollection
            Get
                Return Me.colLicenseeCoursesTest
            End Get

            Set(ByVal value As MUSTER.Info.LicenseeCourseTestCollection)
                Me.colLicenseeCoursesTest = value
            End Set
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        'Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal ID As Integer) As MUSTER.Info.LicenseeCourseTestInfo
            Dim oLicenseeCourseTestInfoLocal As MUSTER.Info.LicenseeCourseTestInfo
            Try
                For Each oLicenseeCourseTestInfoLocal In colLicenseeCoursesTest.Values
                    If oLicenseeCourseTestInfoLocal.ID = ID Then
                        oLicenseeCourseTestInfo = oLicenseeCourseTestInfoLocal
                        Return oLicenseeCourseTestInfo
                    End If
                Next
                oLicenseeCourseTestInfo = oLicenseeCourseTestDB.DBGetByID(ID)
                If oLicenseeCourseTestInfo.ID = 0 Then
                    oLicenseeCourseTestInfo.ID = nID
                    nID -= 1
                End If
                colLicenseeCoursesTest.Add(oLicenseeCourseTestInfo)
                Return oLicenseeCourseTestInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function Retrieve(ByVal LicenseeCoursesTestName As String) As MUSTER.Info.LicenseeCourseTestInfo
            'Try
            '    oLicenseeCourseTestInfo = Nothing
            '    If colLicenseeCoursesTest.Contains(LicenseeCoursesTestName) Then
            '        oLicenseeCourseTestInfo = colLicenseeCoursesTest(LicenseeCoursesTestName)
            '    Else
            '        If oLicenseeCourseTestInfo Is Nothing Then
            '            oLicenseeCourseTestInfo = New MUSTER.Info.LicenseeCourseTestInfo
            '        End If
            '        oLicenseeCourseTestInfo = oLicenseeCoursesTestDB.DBGetByName(LicenseeCoursesTestName)
            '        colLicenseeCoursesTest.Add(oLicenseeCourseTestInfo)
            '    End If
            '    Return oLicenseeCourseTestInfo
            'Catch Ex As Exception
            '    If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
            '    Throw Ex
            'End Try

        End Function
        'Saves the data in the current Info object
        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal bolValidated As Boolean = False)
            Try
                'If Me.ValidateData() Then
                Dim OldKey As String = oLicenseeCourseTestInfo.ID.ToString
                oLicenseeCourseTestDB.Put(oLicenseeCourseTestInfo, moduleID, staffID, returnVal)
                If Not returnVal = String.Empty Then
                    Exit Sub
                End If

                If Not bolValidated Then
                    If oLicenseeCourseTestInfo.ID.ToString <> OldKey Then
                        colLicenseeCoursesTest.ChangeKey(OldKey, oLicenseeCourseTestInfo.ID.ToString)
                    End If
                End If
                oLicenseeCourseTestInfo.Archive()
                oLicenseeCourseTestInfo.IsDirty = False
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
                If oLicenseeCourseTestInfo.ID <> 0 Then
                    If oLicenseeCourseTestInfo.CourseTypeID <> 0 Then
                        If oLicenseeCourseTestInfo.StartTime <> String.Empty Then
                            If Date.Compare(oLicenseeCourseTestInfo.TestDate, CDate("01/01/0001")) = 0 Then
                                errStr += "Test Date cannot be empty" + vbCrLf
                                validateSuccess = False
                            Else
                                validateSuccess = True
                            End If
                        Else
                            errStr += "Start Time cannot be empty" + vbCrLf
                            validateSuccess = False
                        End If
                    Else
                        errStr += "Course Type cannot be empty" + vbCrLf
                        validateSuccess = False
                    End If
                End If

                If errStr.Length > 0 Or Not validateSuccess Then
                    RaiseEvent LicenseeCourseTestsErr(errStr)
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
        Function GetAll(Optional ByVal nLicenseeID As Integer = 0) As MUSTER.Info.LicenseeCourseTestCollection
            Try
                colLicenseeCoursesTest.Clear()
                colLicenseeCoursesTest = oLicenseeCourseTestDB.DBGetByLicenseeID(nLicenseeID)
                Return colLicenseeCoursesTest
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal ID As Integer)
            Try
                Dim oLicCourseTestInfo As MUSTER.Info.LicenseeCourseTestInfo
                If ID = 0 Then
                    oLicCourseTestInfo = New MUSTER.Info.LicenseeCourseTestInfo
                    oLicCourseTestInfo.ID = nID
                    nID -= 1
                    oLicenseeCourseTestInfo = oLicCourseTestInfo
                Else
                    oLicenseeCourseTestInfo = oLicenseeCourseTestDB.DBGetByID(ID)
                End If

                colLicenseeCoursesTest.Add(oLicenseeCourseTestInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oLicenseeCourseTest As MUSTER.Info.LicenseeCourseTestInfo)
            Try
                oLicenseeCourseTestInfo = oLicenseeCourseTest
                colLicenseeCoursesTest.Add(oLicenseeCourseTestInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Integer)
            Dim myIndex As Int16 = 1
            Dim oLicenseeCourseTestInfoLocal As MUSTER.Info.LicenseeCourseTestInfo

            Try
                For Each oLicenseeCourseTestInfoLocal In colLicenseeCoursesTest.Values
                    If oLicenseeCourseTestInfoLocal.ID = ID Then
                        colLicenseeCoursesTest.Remove(oLicenseeCourseTestInfoLocal)
                        Exit Sub
                    End If
                    myIndex += 1
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("LicenseeCourseTest " & ID.ToString & " is not in the collection of LicenseeCourseTest.")
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oLicenseeCourseTest As MUSTER.Info.LicenseeCourseTestInfo)
            Try
                colLicenseeCoursesTest.Remove(oLicenseeCourseTest)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("LicenseeCourseTest " & oLicenseeCourseTest.ID & " is not in the collection of LicenseeCourseTest.")
        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal nLicenseeID As Integer = 0)
            Try
                Dim IDs As New Collection
                Dim index As Integer
                Dim xLicenseeCourseTestInfo As MUSTER.Info.LicenseeCourseTestInfo

                For Each xLicenseeCourseTestInfo In colLicenseeCoursesTest.Values
                    If xLicenseeCourseTestInfo.IsDirty Then
                        If nLicenseeID > 0 Then
                            xLicenseeCourseTestInfo.LicenseeID = nLicenseeID
                        End If
                        oLicenseeCourseTestInfo = xLicenseeCourseTestInfo
                        IDs.Add(oLicenseeCourseTestInfo.ID)
                        Me.Save(moduleID, staffID, returnVal, True)
                    End If
                Next

                If Not (IDs Is Nothing) Then
                    For index = 1 To IDs.Count
                        Dim colKey As String = CType(IDs.Item(index), String)
                        xLicenseeCourseTestInfo = colLicenseeCoursesTest.Item(colKey)
                        colLicenseeCoursesTest.ChangeKey(colKey, xLicenseeCourseTestInfo.ID)
                    Next
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
#End Region
#Region "General Operations"
        Public Sub Clear()
            ' oLicenseeCourseTestsInfo = New MUSTER.Info.LicenseeCourseTestInfo
        End Sub
        Public Sub Reset()
            oLicenseeCourseTestInfo.Reset()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function TestTable() As DataTable
            Dim oLicenseeCourseTestsInfoLocal As New MUSTER.Info.LicenseeCourseTestInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Try

                tbEntityTable.Columns.Add("TestID", GetType(Integer))
                tbEntityTable.Columns.Add("LicenseeID", GetType(Integer))
                tbEntityTable.Columns("LicenseeID").DefaultValue = 0
                tbEntityTable.Columns.Add("Date", GetType(Date))
                tbEntityTable.Columns("Date").DefaultValue = Today.ToShortDateString
                tbEntityTable.Columns.Add("StartTime")
                tbEntityTable.Columns.Add("Type")
                tbEntityTable.Columns.Add("Score", GetType(String))
                tbEntityTable.Columns.Add("Deleted", GetType(Boolean))
                tbEntityTable.Columns("Deleted").DefaultValue = False
                tbEntityTable.Columns.Add("Created By")
                tbEntityTable.Columns.Add("Date Created")
                tbEntityTable.Columns.Add("Last Edited By")
                tbEntityTable.Columns.Add("Date Last Edited")


                For Each oLicenseeCourseTestsInfoLocal In colLicenseeCoursesTest.Values
                    dr = tbEntityTable.NewRow()
                    dr("TestID") = oLicenseeCourseTestsInfoLocal.ID
                    dr("LicenseeID") = oLicenseeCourseTestsInfoLocal.LicenseeID
                    dr("Date") = oLicenseeCourseTestsInfoLocal.TestDate
                    dr("StartTime") = oLicenseeCourseTestsInfoLocal.StartTime
                    dr("Type") = oLicenseeCourseTestsInfoLocal.CourseTypeID
                    dr("Score") = oLicenseeCourseTestsInfoLocal.TestScore
                    dr("Deleted") = oLicenseeCourseTestsInfoLocal.Deleted
                    dr("Created By") = oLicenseeCourseTestsInfoLocal.CreatedBy
                    dr("Date Created") = oLicenseeCourseTestsInfoLocal.CreatedOn
                    dr("Last Edited By") = oLicenseeCourseTestsInfoLocal.ModifiedBy
                    dr("Date Last Edited") = oLicenseeCourseTestsInfoLocal.ModifiedOn
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
        Public Function ListStartTime() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vCOM_COURSETEST_STARTTIME")
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

                dsReturn = oLicenseeCourseTestDB.DBGetDS(strSQL)
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
        Private Sub pLicenseeCourseTestsInfoChanged(ByVal bolValue As Boolean) Handles oLicenseeCourseTestInfo.LicenseeCourseTestInfoChanged
            RaiseEvent LicenseeCourseTestsChanged(bolValue)
        End Sub
        Private Sub pLicenseeCourseTestsColChanged(ByVal bolValue As Boolean) Handles colLicenseeCoursesTest.LicenseeCourseTestColChanged
            RaiseEvent ColChanged(bolValue)
        End Sub
#End Region
    End Class
End Namespace
