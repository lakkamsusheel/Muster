'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.Persona
'   Provides the info and collection objects to the client for manipulating
'   an AddressInfo object.
'   Copyright (C) 2004 CIBER, Inc.
'  All rights reserved.
'
' Release   Initials    Date        Description
'   1.0         EN      12/13/04    Original class definition.
'   1.1         EN      12/24/04    Added the properties in to collection while setting the property.
'   1.2         AN      01/03/05    Added Try catch and Exception Handling/Logging
'   1.3         EN      01/06/05    Modified Reset Method.
'   1.4         EN      01/10/05    Added  oPersonaInfo.Archive() in save method.
'   1.5         MNR     01/13/05    Added  Events
'   1.6         MNR     01/14/05    Added ValidateData(), modified flush(), modified Remove(ID),
'                                   Replaced events PersonaErr with evtPersonaErr
'   1.7         EN      01/18/05    Modified Data type in Organization Entity Code to Integer.
'   1.8         EN      01/21/05    Added Look up Functions. 
'   1.9         MNR     01/25/05    Added GetNext(), GetPrevious() and GetNextPrev(..) functions
'   1.8         JVC2    02/02/05    Added EntityTypeID to private members and initialize to "Persona" type.
'                                       Added second optional parameter to SAVE to specify module.
'                                       Added second optional module to ValidateData for forcing validation
'   1.9         MNR     03/16/05    Removed strSrc from events
'
' Function          Description
' Retrieve(ID)     Returns the person/Org requested by the int arg ID
' GetAll()    Returns an  PersonaCollection with all Person objects
' Add(ID)           Adds the Persona identified by arg ID to the 
'                           internal  PersonaCollection
' Add(Name)         Adds the Person identified by arg NAME to the internal 
'                    PersonaCollection            
' Add(Person)       Adds the Person passed as the argument to the internal 
'                            PersonaCollection
' Remove(ID)        Removes the Person identified by arg ID from the internal 
'                            PersonaCollection
' Remove(NAME)      Removes the Person identified by arg NAME from the 
'                           internal  PersonaCollection
' PersonTable()     Returns a datatable containing all columns for the Person 

' colIsDirty()       Returns a boolean indicating whether any of the PersonaInfo
'                    objects in the PersonaCollection has been modified since the
'                    last time it was retrieved from/saved to the repository.
' Flush()            Marshalls all modified/added PersonaInfo objects in the 
'                        PersonaCollection to the repository.
' Save()             Marshalls the internal PersonaInfo object to the repository.

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pPersona
#Region "Public Events"
        Public Event PersonaErr(ByVal MsgStr As String, ByVal strColumnName As String)
        Public Event evtPersonaErr(ByVal MsgStr As String)
        Public Event evtPersonaChanged(ByVal bolValue As Boolean)
        Public Event evtPersonasChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        Private WithEvents colPersona As Muster.Info.PersonaCollection
        Private WithEvents oPersonaInfo As Muster.Info.PersonaInfo
        Private opersonaDB As New Muster.DataAccess.PersonaDB
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private strNewID As String
        Private nID As Int64 = -1
        Private bolShowDeleted As Boolean
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("Persona").ID
#End Region
#Region "Constructors"
        Public Sub New()
            oPersonaInfo = New Muster.Info.PersonaInfo
            colPersona = New Muster.Info.PersonaCollection
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As String
            Get
                Return oPersonaInfo.ID
            End Get
            Set(ByVal Value As String)
                oPersonaInfo.ID = Value
                'colPersona.Add(oPersonaInfo)
            End Set
        End Property
        Public Property PersonId() As Integer
            Get
                Return oPersonaInfo.PersonId
            End Get

            Set(ByVal value As Integer)
                oPersonaInfo.PersonId = Integer.Parse(value)
                colPersona(oPersonaInfo.ID) = oPersonaInfo
            End Set
        End Property
        Public Property OrgID() As Integer
            Get
                Return oPersonaInfo.OrgID
            End Get

            Set(ByVal value As Integer)
                oPersonaInfo.OrgID = Integer.Parse(value)
                colPersona(oPersonaInfo.ID) = oPersonaInfo
            End Set
        End Property
        Public Property Org_Entity_Code() As Integer
            Get
                Return oPersonaInfo.Org_Entity_Code
            End Get
            Set(ByVal value As Integer)
                If oPersonaInfo.PersonId <> 0 Then
                    'RaiseEvent PersonaErr("You cannot Enter Organization details for Person.", "ORGANIZATION_ENTITY_CODE")
                    RaiseEvent evtPersonaErr("You cannot Enter Organization details for Person.")
                    Exit Property
                Else
                    oPersonaInfo.Org_Entity_Code = value
                    colPersona(oPersonaInfo.ID) = oPersonaInfo
                End If

            End Set
        End Property
        Public Property Company() As String
            Get
                Return oPersonaInfo.Company
            End Get

            Set(ByVal value As String)
                If oPersonaInfo.PersonId <> 0 Then
                    'RaiseEvent PersonaErr("You cannot Enter Organization details for Person.", "NAME")
                    RaiseEvent evtPersonaErr("You cannot Enter Organization details for Person.")
                    Exit Property
                Else
                    oPersonaInfo.Company = value
                    colPersona(oPersonaInfo.ID) = oPersonaInfo
                End If
            End Set
        End Property
        Public Property Title() As String
            Get
                Return oPersonaInfo.Title
            End Get
            Set(ByVal value As String)
                If oPersonaInfo.OrgID <> 0 Then
                    'RaiseEvent PersonaErr("You cannot Enter Person details for Organization.", "TITLE")
                    RaiseEvent evtPersonaErr("You cannot Enter Person details for Organization.")
                    Exit Property
                Else
                    oPersonaInfo.Title = value
                    colPersona(oPersonaInfo.ID) = oPersonaInfo
                End If
            End Set
        End Property
        Public Property Prefix() As String
            Get
                Return oPersonaInfo.Prefix
            End Get

            Set(ByVal value As String)
                If oPersonaInfo.OrgID <> 0 Then
                    'RaiseEvent PersonaErr("You cannot Enter Prefix Information for Organization.", "PREFIX")
                    RaiseEvent evtPersonaErr("You cannot Enter Prefix Information for Organization.")
                    Exit Property
                Else
                    oPersonaInfo.Prefix = value
                    colPersona(oPersonaInfo.ID) = oPersonaInfo
                End If
            End Set
        End Property
        Public Property FirstName() As String
            Get
                Return oPersonaInfo.FirstName
            End Get
            Set(ByVal value As String)
                If oPersonaInfo.OrgID <> 0 Then
                    'RaiseEvent PersonaErr("You cannot Enter FirstName Information for Organization.", "FIRST_NAME")
                    RaiseEvent evtPersonaErr("You cannot Enter FirstName Information for Organization.")
                    Exit Property
                Else
                    oPersonaInfo.FirstName = value
                    colPersona(oPersonaInfo.ID) = oPersonaInfo
                End If
            End Set
        End Property
        Public Property MiddleName() As String
            Get
                Return oPersonaInfo.MiddleName
            End Get
            Set(ByVal value As String)
                If oPersonaInfo.OrgID <> 0 Then
                    'RaiseEvent PersonaErr("You cannot Enter MiddleName Information for Organization.", "MIDDLE_NAME")
                    RaiseEvent evtPersonaErr("You cannot Enter MiddleName Information for Organization.")
                    Exit Property
                Else
                    oPersonaInfo.MiddleName = value
                    colPersona(oPersonaInfo.ID) = oPersonaInfo
                End If
            End Set
        End Property
        Public Property LastName() As String
            Get
                Return oPersonaInfo.LastName
            End Get
            Set(ByVal value As String)
                If oPersonaInfo.OrgID <> 0 Then
                    'RaiseEvent PersonaErr("You cannot Enter LastName Information for Organization.", "LAST_NAME")
                    RaiseEvent evtPersonaErr("You cannot Enter LastName Information for Organization.")
                    Exit Property
                Else
                    oPersonaInfo.LastName = value
                    colPersona(oPersonaInfo.ID) = oPersonaInfo
                End If
            End Set
        End Property
        Public Property Suffix() As String
            Get
                Return oPersonaInfo.Suffix
            End Get

            Set(ByVal value As String)
                If oPersonaInfo.OrgID <> 0 Then
                    'RaiseEvent PersonaErr("You cannot Enter Suffix Information for Organization.", "SUFFIX")
                    RaiseEvent evtPersonaErr("You cannot Enter Suffix Information for Organization.")
                    Exit Property
                Else
                    oPersonaInfo.Suffix = value
                    colPersona(oPersonaInfo.ID) = oPersonaInfo
                End If
            End Set
        End Property
        'Public ReadOnly Property EntityType() As Integer
        '    Get
        '        Return nEntityTypeID
        '    End Get
        'End Property
        Public Property Deleted() As Boolean
            Get
                Return oPersonaInfo.Deleted
            End Get

            Set(ByVal value As Boolean)
                oPersonaInfo.Deleted = value
                colPersona(oPersonaInfo.ID) = oPersonaInfo
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oPersonaInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oPersonaInfo.IsDirty = value
                colPersona(oPersonaInfo.ID) = oPersonaInfo
            End Set
        End Property
        Public Property CreatedBy() As String
            Get
                Return oPersonaInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oPersonaInfo.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oPersonaInfo.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oPersonaInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oPersonaInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oPersonaInfo.ModifiedOn
            End Get
        End Property
        Public Property colIsDirty() As Boolean
            Get
                Dim xPersonainfo As MUSTER.Info.PersonaInfo
                For Each xPersonainfo In colPersona.Values
                    If xPersonainfo.IsDirty Then
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
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        Public Sub Load(ByVal ds As DataSet, ByVal personaType As String)
            Dim dr As DataRow
            Try
                For Each dr In ds.Tables("OrgPerson").Rows
                    oPersonaInfo = New MUSTER.Info.PersonaInfo(dr, personaType)
                    colPersona.Add(oPersonaInfo)
                Next
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal bolValidated As Boolean = False) As Boolean
            Dim bolSaveSuccess As Boolean = True
            Try
                If Not bolValidated Then
                    If Not Me.ValidateData(moduleID, True) Then
                        bolSaveSuccess = False
                        Exit Try
                    End If
                End If
                strNewID = opersonaDB.Put(oPersonaInfo, moduleID, staffID, returnVal)
                If Not returnVal = String.Empty Then
                    Exit Function
                End If

                If strNewID <> "" Then
                    oPersonaInfo.ID = strNewID
                    If strNewID.Chars(0) = "O" Then
                        oPersonaInfo.OrgID = strNewID.Substring(2)
                    ElseIf strNewID.Chars(0) = "P" Then
                        oPersonaInfo.PersonId = (strNewID.Substring(2))
                    End If
                End If
                oPersonaInfo.Archive()
                oPersonaInfo.IsDirty = False
                RaiseEvent evtPersonaChanged(oPersonaInfo.IsDirty)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Return bolSaveSuccess
        End Function
        Public Function Retrieve(ByVal FullKey As String, Optional ByVal ShowDeleted As Boolean = False) As MUSTER.Info.PersonaInfo
            Dim bolDataAged As Boolean = False
            Try

                'Dim keyID As Integer = 0
                'If Not (oPersonaInfo.ID Is Nothing Or oPersonaInfo.ID = String.Empty) Then
                'keyID = CType(oPersonaInfo.ID.Split("|")(1), Integer)
                'End If
                'If keyID < 0 And _
                '   oPersonaInfo.FirstName = String.Empty And _
                '  oPersonaInfo.LastName = String.Empty And _
                ' oPersonaInfo.Company = String.Empty And _
                '   oPersonaInfo.Org_Entity_Code = 0 And _
                '  Not oPersonaInfo.IsDirty And _
                '  FullKey = "P|0" Then
                ' Exit Try
                '  End If
                'Me.ValidateData()

                If colPersona.Contains(FullKey) Then
                    oPersonaInfo = colPersona.Item(FullKey)
                    If oPersonaInfo.IsAgedData = True And oPersonaInfo.IsDirty = False Then
                        bolDataAged = True
                    Else
                        Return oPersonaInfo
                    End If
                End If
                If bolDataAged Then
                    colPersona.Remove(oPersonaInfo)
                End If

                Dim strArray() As String
                strArray = FullKey.Split("|")
                oPersonaInfo = opersonaDB.DBGetByKey(strArray, ShowDeleted)
                If oPersonaInfo.PersonId = 0 And oPersonaInfo.OrgID = 0 Then
                    oPersonaInfo.ID = 0 & "|" & nID
                    nID -= 1
                End If
                colPersona.Add(oPersonaInfo)

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Return oPersonaInfo
        End Function
        Public Function ValidateData(Optional ByVal moduleID As Integer = 612, Optional ByVal ForceValidation As Boolean = False) As Boolean
            ' to be completed according to DDD specs for registration / technical ... - Manju 01/14/05
            Try
                Dim keyID As Integer
                Dim errStr As String = ""
                Dim validateSuccess As Boolean = True
                Select Case moduleID
                    Case 612        'Registration
                        If oPersonaInfo.ID <> String.Empty Or ForceValidation Then
                            '
                            ' Replaced keyID with oPersonaInfo.Ambi_ID
                            If oPersonaInfo.Ambi_ID < 0 And _
                                oPersonaInfo.FirstName = String.Empty And _
                                oPersonaInfo.LastName = String.Empty And _
                                oPersonaInfo.Company = String.Empty And _
                                oPersonaInfo.Org_Entity_Code = 0 And _
                                oPersonaInfo.IsDirty Then
                                errStr += "Required Fields cannot be empty" + vbCrLf
                                validateSuccess = False
                                Exit Select
                            End If
                            If oPersonaInfo.PersonId = 0 And oPersonaInfo.OrgID <> 0 Then
                                ' if organization
                                If oPersonaInfo.Org_Entity_Code <> 0 Then
                                    If oPersonaInfo.Company <> String.Empty Then
                                        validateSuccess = True
                                        Exit Select
                                    Else
                                        errStr += "Company cannot be empty" + vbCrLf
                                        validateSuccess = True
                                    End If
                                Else
                                    oPersonaInfo.Org_Entity_Code = 539

                                    validateSuccess = True
                                End If
                            ElseIf oPersonaInfo.PersonId <> 0 And oPersonaInfo.OrgID = 0 Then
                                ' if person
                                If oPersonaInfo.FirstName <> String.Empty Then
                                    If oPersonaInfo.LastName <> String.Empty Then
                                        validateSuccess = True
                                    Else
                                        errStr += "Last Name cannot be empty" + vbCrLf
                                        validateSuccess = False
                                    End If
                                Else
                                    errStr += "First Name cannot be empty" + vbCrLf
                                    validateSuccess = False
                                End If
                            Else
                                ' if both person / org id = 0
                                If oPersonaInfo.FirstName <> String.Empty And _
                                    oPersonaInfo.LastName <> String.Empty Then
                                    validateSuccess = True
                                    Exit Select
                                ElseIf oPersonaInfo.Company <> String.Empty And _
                                        oPersonaInfo.Org_Entity_Code <> 0 Then
                                    validateSuccess = True
                                    Exit Select
                                Else
                                    If (oPersonaInfo.FirstName = String.Empty Or _
                                        oPersonaInfo.LastName = String.Empty) And oPersonaInfo.Company = String.Empty Then
                                        errStr += "(First and Last Name) or company are required Fields" + vbCrLf
                                        validateSuccess = False
                                    ElseIf oPersonaInfo.Org_Entity_Code = 0 AndAlso oPersonaInfo.Company <> String.Empty Then
                                        oPersonaInfo.Org_Entity_Code = 539
                                        validateSuccess = True
                                    Else
                                        errStr += "Persona has to be either Person or Organization" + vbCrLf
                                        validateSuccess = False
                                    End If
                                End If
                            End If
                        End If
                        Exit Select
                        'Case "Technical"
                End Select
                If errStr.Length > 0 And Not validateSuccess Then
                    RaiseEvent evtPersonaErr(errStr)
                End If
                Return validateSuccess
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region
#Region "Collection Operations"
        Function GetAll(Optional ByVal ShowDeleted As Boolean = False) As MUSTER.Info.PersonaCollection
            Try
                colPersona.Clear()
                colPersona = opersonaDB.GetAllInfo(ShowDeleted)
                Return colPersona
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Sub Add(ByRef oPersona As MUSTER.Info.PersonaInfo)
            Try
                oPersonaInfo = oPersona
                If oPersonaInfo.ID = Nothing Then
                    oPersonaInfo.ID = 0 & "|" & nID
                    nID -= 1
                End If
                colPersona.Add(oPersonaInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        'Removes the address called for by ID from the collection
        Public Sub Remove(ByVal ID As String)
            Dim myIndex As Int16 = 1
            Dim oPersonaInfoLocal As MUSTER.Info.PersonaInfo
            Try
                oPersonaInfoLocal = colPersona.Item(ID)
                If Not (oPersonaInfoLocal Is Nothing) Then
                    colPersona.Remove(oPersonaInfoLocal)
                    Exit Sub
                End If
                'For Each oPersonaInfoLocal In colPersona.Values
                '    If oPersonaInfoLocal.ID = ID Then
                '        colPersona.Remove(oPersonaInfoLocal)
                '        'oPersonaInfo = New Muster.Info.PersonaInfo
                '        Exit Sub
                '    End If
                '    myIndex += 1
                'Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

            Throw New Exception("Persona " & ID.ToString & " is not in the collection of Persona.")
        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try
                Dim IDs As New Collection
                Dim index As Integer
                Dim IDKey As Integer
                Dim oTempInfo As MUSTER.Info.PersonaInfo
                For Each oTempInfo In colPersona.Values
                    If oTempInfo.IsDirty Then
                        oPersonaInfo = oTempInfo
                        If Me.ValidateData() Then
                            IDKey = oPersonaInfo.ID.Split("|")(1)
                            If IDKey < 0 And _
                                Not oPersonaInfo.Deleted Then
                                IDs.Add(oPersonaInfo.ID)
                            End If
                            Me.Save(moduleID, staffID, returnVal, True)
                        Else : Exit For
                        End If
                    End If
                Next
                If Not (IDs Is Nothing) Then
                    For index = 1 To IDs.Count
                        Dim colKey As String = CType(IDs.Item(index), String)
                        oTempInfo = colPersona.Item(colKey)
                        colPersona.ChangeKey(colKey, oTempInfo.ID)
                    Next
                End If
                RaiseEvent evtPersonasChanged(Me.colIsDirty)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function GetNext() As String
            Return GetNextPrev(1)
        End Function
        Public Function GetPrevious() As String
            Return GetNextPrev(-1)
        End Function
        Private Function GetNextPrev(ByVal direction As Integer) As String
            Dim strArr() As String = colPersona.GetKeys()
            'Dim nArr(strArr.GetUpperBound(0)) As Integer
            'Dim y As String
            'For Each y In strArr
            '    nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            'Next
            'nArr.Sort(nArr)
            colIndex = Array.BinarySearch(strArr, Me.ID.ToString)
            If colIndex + direction > -1 And _
                colIndex + direction <= strArr.GetUpperBound(0) Then
                Return colPersona.Item(strArr.GetValue(colIndex + direction)).ID.ToString
            Else
                Return colPersona.Item(strArr.GetValue(colIndex)).ID.ToString
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear(Optional ByVal strDepth As String = "ALL")
            'oPersonaInfo = New Muster.Info.PersonaInfo
            Reset()
            oPersonaInfo = Retrieve("P|0")
        End Sub
        Public Sub Reset(Optional ByVal strDepth As String = "ALL")
            oPersonaInfo.Reset()
        End Sub
#End Region
#Region "Look Up Operations"
        Public Function PopulateEntityCode() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vORG_ENTITY_TYPE")
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
                dsReturn = opersonaDB.DBGetDS(strSQL)
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
        'Returns a datatable of the addresses in the collection
        Public Function PersonaTable() As DataTable

            Dim opersonaInfoLocal As MUSTER.Info.PersonaInfo
            Dim dr As DataRow
            Dim tbPersonaTable As New DataTable

            Try
                tbPersonaTable.Columns.Add("ID")
                tbPersonaTable.Columns.Add("Organization_Id")
                tbPersonaTable.Columns.Add("Person_ID")
                tbPersonaTable.Columns.Add("CompanyName")
                tbPersonaTable.Columns.Add("Title")
                tbPersonaTable.Columns.Add("Prefix")
                tbPersonaTable.Columns.Add("First_name")
                tbPersonaTable.Columns.Add("Middle_name")
                tbPersonaTable.Columns.Add("Last_name")
                tbPersonaTable.Columns.Add("Suffix")
                tbPersonaTable.Columns.Add("organization_entity_code")
                tbPersonaTable.Columns.Add("DELETED")
                tbPersonaTable.Columns.Add("created_by")
                tbPersonaTable.Columns.Add("date_created")
                tbPersonaTable.Columns.Add("last_edited_by")
                tbPersonaTable.Columns.Add("date_last_edited")

                For Each opersonaInfoLocal In colPersona.Values
                    dr = tbPersonaTable.NewRow()
                    dr("ID") = opersonaInfoLocal.ID
                    dr("Organization_Id") = opersonaInfoLocal.OrgID
                    dr("Person_ID") = opersonaInfoLocal.PersonId
                    dr("CompanyName") = opersonaInfoLocal.Company
                    dr("Title") = opersonaInfoLocal.Title
                    dr("Prefix") = opersonaInfoLocal.Prefix
                    dr("First_name") = opersonaInfoLocal.FirstName
                    dr("Middle_name") = opersonaInfoLocal.MiddleName
                    dr("Last_name") = opersonaInfoLocal.LastName
                    dr("Suffix") = opersonaInfoLocal.Suffix
                    dr("organization_entity_code") = opersonaInfoLocal.Org_Entity_Code
                    dr("DELETED") = opersonaInfoLocal.Deleted
                    dr("CREATED_BY") = opersonaInfoLocal.CreatedBy
                    dr("DATE_CREATED") = opersonaInfoLocal.CreatedOn
                    dr("LAST_EDITED_BY") = opersonaInfoLocal.ModifiedBy
                    dr("DATE_LAST_EDITED") = opersonaInfoLocal.ModifiedOn
                    tbPersonaTable.Rows.Add(dr)
                Next
                Return tbPersonaTable
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Function GetDataSet(ByVal strSQL As String) As DataSet
            Try
                Dim ds As DataSet
                ds = opersonaDB.DBGetDS(strSQL)
                Return ds
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
#End Region
#End Region
#Region "Event Handlers"
        Private Sub PersonasChanged(ByVal strSrc As String) Handles colPersona.PersonaColChanged
            RaiseEvent evtPersonasChanged(Me.colIsDirty)
        End Sub
        Private Sub PersonaChanged(ByVal bolValue As Boolean) Handles oPersonaInfo.PersonaInfoChanged
            RaiseEvent evtPersonaChanged(bolValue)
        End Sub

        Public Sub ManuallyAlertChange()
            RaiseEvent evtPersonaChanged(True)
        End Sub
#End Region
    End Class
End Namespace
