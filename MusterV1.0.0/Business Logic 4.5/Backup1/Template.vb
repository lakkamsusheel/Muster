'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.Template
'   Provides the operations required to manipulate an Entity object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0                              Original class definition
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
' NOTE: This file to be used as Template to build other objects.
'       Replace keyword "Template" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pTemplate
#Region "Public Events"
        Public Event TemplateErr(ByVal MsgStr As String)
        Public Event TemplateChanged(ByVal bolValue As Boolean)
        Public Event ColChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        Private WithEvents oTemplateInfo As Muster.Info.TemplateInfo
        Private WithEvents colTemplates As Muster.Info.TemplatesCollection
        Private oTemplateDB As New Muster.DataAccess.TemplateDB
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private nID As Int64 = -1
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
#End Region
#Region "Constructors"
        Public Sub New()
            oTemplateInfo = New Muster.Info.TemplateInfo
            colTemplates = New Muster.Info.TemplatesCollection
        End Sub
        '********************************************************
        '
        ' Overloaded NEW which will populate with a single instance
        '   of the named Template object.
        '
        '********************************************************
        Public Sub New(ByVal TemplateName As String)
            oTemplateInfo = New Muster.Info.TemplateInfo
            colTemplates = New Muster.Info.TemplatesCollection
            Me.Retrieve(TemplateName)
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Integer
            Get
                Return oTemplateInfo.ID
            End Get
            Set(ByVal Value As Integer)
                oTemplateInfo.ID = Integer.Parse(Value)
            End Set
        End Property
        Public Property Name() As String
            Get
                Return oTemplateInfo.Name
            End Get
            Set(ByVal Value As String)
                oTemplateInfo.Name = Value
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oTemplateInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oTemplateInfo.Deleted = Boolean.Parse(Value)
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oTemplateInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oTemplateInfo.IsDirty = Boolean.Parse(value)
            End Set
        End Property
        Public Property colIsDirty() As Boolean
            Get
                Dim xTemplateinfo As Muster.Info.TemplateInfo
                For Each xTemplateinfo In colTemplates.Values
                    If xTemplateinfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
            Set(ByVal Value As Boolean)
                oTemplateInfo.IsDirty = Value
            End Set
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        'Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal ID As Integer) As Muster.Info.TemplateInfo
            Dim oTemplateInfoLocal As Muster.Info.TemplateInfo
            Try
                For Each oTemplateInfoLocal In colTemplates.Values
                    If oTemplateInfoLocal.ID = ID Then
                        oTemplateInfo = oTemplateInfoLocal
                        Return oTemplateInfo
                    End If
                Next
                oTemplateInfo = oTemplateDB.DBGetByID(ID)
                If oTemplateInfo.ID = 0 Then
                    oTemplateInfo.ID = nID
                    nID -= 1
                End If
                colTemplates.Add(oTemplateInfo)
                Return oTemplateInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function Retrieve(ByVal TemplateName As String) As Muster.Info.TemplateInfo
            Try
                oTemplateInfo = Nothing
                If colTemplates.Contains(TemplateName) Then
                    oTemplateInfo = colTemplates(TemplateName)
                Else
                    If oTemplateInfo Is Nothing Then
                        oTemplateInfo = New Muster.Info.TemplateInfo
                    End If
                    oTemplateInfo = oTemplateDB.DBGetByName(TemplateName)
                    colTemplates.Add(oTemplateInfo)
                End If
                Return oTemplateInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        'Saves the data in the current Info object
        Public Sub Save()
            Dim strModuleName As String = String.Empty
            Try
                If Me.ValidateData(strModuleName) Then
                    oTemplateDB.Put(oTemplateInfo)
                    oTemplateInfo.Archive()
                End If
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
        Function GetAll() As Muster.Info.TemplatesCollection
            Try
                colTemplates.Clear()
                colTemplates = oTemplateDB.GetAllInfo
                Return colTemplates
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal ID As Integer)
            Try
                oTemplateInfo = oTemplateDB.DBGetByID(ID)
                If oTemplateInfo.ID = 0 Then
                    oTemplateInfo.ID = nID
                    nID -= 1
                End If
                colTemplates.Add(oTemplateInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oTemplate As Muster.Info.TemplateInfo)
            Try
                oTemplateInfo = oTemplate
                colTemplates.Add(oTemplateInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Integer)
            Dim myIndex As Int16 = 1
            Dim oTemplateInfoLocal As Muster.Info.TemplateInfo

            Try
                For Each oTemplateInfoLocal In colTemplates.Values
                    If oTemplateInfoLocal.ID = ID Then
                        colTemplates.Remove(oTemplateInfoLocal)
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
        Public Sub Remove(ByVal oTemplate As Muster.Info.TemplateInfo)
            Try
                colTemplates.Remove(oTemplate)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("Template " & oTemplate.ID & " is not in the collection of Templates.")
        End Sub
        Public Sub Flush()
            Dim xTemplateInfo As Muster.Info.TemplateInfo
            For Each xTemplateInfo In colTemplates.Values
                If xTemplateInfo.IsDirty Then
                    oTemplateInfo = xTemplateInfo
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
                    Me.Save()
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
            Dim strArr() As String = colTemplates.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colTemplates.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            Else
                Return colTemplates.Item(nArr.GetValue(colIndex)).ID.ToString
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oTemplateInfo = New MUSTER.Info.TemplateInfo
        End Sub
        Public Sub Reset()
            oTemplateInfo.Reset()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function EntityTable() As DataTable
            Dim oTemplateInfoLocal As New Muster.Info.PipeInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Try
                tbEntityTable.Columns.Add("Template ID")
                tbEntityTable.Columns.Add("Deleted")
                tbEntityTable.Columns.Add("Created By")
                tbEntityTable.Columns.Add("Date Created")
                tbEntityTable.Columns.Add("Last Edited By")
                tbEntityTable.Columns.Add("Date Last Edited")

                For Each oTemplateInfoLocal In colTemplates.Values
                    dr = tbEntityTable.NewRow()
                    dr("Pipe ID") = oTemplateInfoLocal.ID
                    dr("Deleted") = oTemplateInfoLocal.Deleted
                    dr("Created By") = oTemplateInfoLocal.CreatedBy
                    dr("Date Created") = oTemplateInfoLocal.CreatedOn
                    dr("Last Edited By") = oTemplateInfoLocal.ModifiedBy
                    dr("Date Last Edited") = oTemplateInfoLocal.ModifiedOn
                    tbEntityTable.Rows.Add(dr)
                Next
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region
#End Region
#Region "External Event Handlers"
        Private Sub TemplateInfoChanged(ByVal bolValue As Boolean) Handles oTemplateInfo.TemplateInfoChanged
            RaiseEvent TemplateChanged(bolValue)
        End Sub
        Private Sub TamplateColChanged(ByVal bolValue As Boolean) Handles colTemplates.TemplateColChanged
            RaiseEvent ColChanged(bolValue)
        End Sub
#End Region
    End Class
End Namespace
