'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.CAEFceCitations
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
    Public Class pCAEFceCitations
#Region "Public Events"
#End Region
#Region "Private Member Variables"
        Private WithEvents oCAEFceCitationInfo As MUSTER.Info.CAEFceCitationInfo
        Private WithEvents colCAEFceCitation As MUSTER.Info.CAEFceCitationCollection
        Private oCAEFceCitationDB As New MUSTER.DataAccess.CAEFceCitationDB
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private nID As Int64 = -1
#End Region
#Region "Constructors"
        Public Sub New()
            oCAEFceCitationInfo = New MUSTER.Info.CAEFceCitationInfo
            colCAEFceCitation = New MUSTER.Info.CAEFceCitationCollection
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Integer
            Get
                Return oCAEFceCitationInfo.ID
            End Get
            Set(ByVal Value As Integer)
                oCAEFceCitationInfo.ID = Integer.Parse(Value)
            End Set
        End Property

        Public Property Deleted() As Boolean
            Get
                Return oCAEFceCitationInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oCAEFceCitationInfo.Deleted = Boolean.Parse(Value)
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oCAEFceCitationInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oCAEFceCitationInfo.IsDirty = Boolean.Parse(value)
            End Set
        End Property
        Public Property colIsDirty() As Boolean
            Get
                Dim xCAEFceCitationInfo As MUSTER.Info.CAEFceCitationInfo
                For Each xCAEFceCitationInfo In colCAEFceCitation.Values
                    If xCAEFceCitationInfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
            Set(ByVal Value As Boolean)
                oCAEFceCitationInfo.IsDirty = Value
            End Set
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        'Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal ID As Integer) As MUSTER.Info.CAEFceCitationInfo
            Dim oCAEFceCitationInfoLocal As MUSTER.Info.CAEFceCitationInfo
            Try
                For Each oCAEFceCitationInfoLocal In colCAEFceCitation.Values
                    If oCAEFceCitationInfoLocal.ID = ID Then
                        oCAEFceCitationInfo = oCAEFceCitationInfoLocal
                        Return oCAEFceCitationInfo
                    End If
                Next
                oCAEFceCitationInfo = oCAEFceCitationDB.DBGetByID(ID)
                If oCAEFceCitationInfo.ID = 0 Then
                    oCAEFceCitationInfo.ID = nID
                    nID -= 1
                End If
                colCAEFceCitation.Add(oCAEFceCitationInfo)
                Return oCAEFceCitationInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function Retrieve(ByVal TemplateName As String) As MUSTER.Info.CAEFceCitationInfo
            Try
                oCAEFceCitationInfo = Nothing
                If colCAEFceCitation.Contains(TemplateName) Then
                    oCAEFceCitationInfo = colCAEFceCitation(TemplateName)
                Else
                    If oCAEFceCitationInfo Is Nothing Then
                        oCAEFceCitationInfo = New MUSTER.Info.CAEFceCitationInfo
                    End If
                    colCAEFceCitation.Add(oCAEFceCitationInfo)
                End If
                Return oCAEFceCitationInfo
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
                    oCAEFceCitationDB.Put(oCAEFceCitationInfo)
                    oCAEFceCitationInfo.Archive()
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
                    'RaiseEvent TemplateErr(errStr)
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
        Function GetAll() As MUSTER.Info.CAEFceCitationCollection
            Try
                colCAEFceCitation.Clear()
                colCAEFceCitation = oCAEFceCitationDB.GetAllInfo
                Return colCAEFceCitation
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal ID As Integer)
            Try
                oCAEFceCitationInfo = oCAEFceCitationDB.DBGetByID(ID)
                If oCAEFceCitationInfo.ID = 0 Then
                    oCAEFceCitationInfo.ID = nID
                    nID -= 1
                End If
                colCAEFceCitation.Add(oCAEFceCitationInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oTemplate As MUSTER.Info.CAEFceCitationInfo)
            Try
                oCAEFceCitationInfo = oTemplate
                colCAEFceCitation.Add(oCAEFceCitationInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Integer)
            Dim myIndex As Int16 = 1
            Dim oCAEFceCitationInfoLocal As MUSTER.Info.CAEFceCitationInfo
            Try
                For Each oCAEFceCitationInfoLocal In colCAEFceCitation.Values
                    If oCAEFceCitationInfoLocal.ID = ID Then
                        colCAEFceCitation.Remove(oCAEFceCitationInfoLocal)
                        Exit Sub
                    End If
                    myIndex += 1
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("Citation " & ID.ToString & " is not in the collection of FCE Citations.")
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oCAEFceCitation As MUSTER.Info.CAEFceCitationInfo)
            Try
                colCAEFceCitation.Remove(oCAEFceCitation)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("FCE Citation " & oCAEFceCitation.ID & " is not in the collection of CAE Fce Citation.")
        End Sub
        Public Sub Flush()
            Dim xCAEFceCitationInfo As MUSTER.Info.CAEFceCitationInfo
            For Each xCAEFceCitationInfo In colCAEFceCitation.Values
                If xCAEFceCitationInfo.IsDirty Then
                    oCAEFceCitationInfo = xCAEFceCitationInfo
                    Me.Save()
                End If
            Next
        End Sub
#End Region
#Region "General Operations"
        Public Function Clear()
            oCAEFceCitationInfo = New MUSTER.Info.CAEFceCitationInfo
        End Function
        Public Function Reset()
            oCAEFceCitationInfo.Reset()
        End Function
#End Region
#Region "Miscellaneous Operations"
        Public Function EntityTable() As DataTable
            Dim oCAEFceCitationInfoLocal As New MUSTER.Info.PipeInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Try
                tbEntityTable.Columns.Add("Template ID")
                tbEntityTable.Columns.Add("Deleted")
                tbEntityTable.Columns.Add("Created By")
                tbEntityTable.Columns.Add("Date Created")
                tbEntityTable.Columns.Add("Last Edited By")
                tbEntityTable.Columns.Add("Date Last Edited")

                For Each oCAEFceCitationInfoLocal In colCAEFceCitation.Values
                    dr = tbEntityTable.NewRow()
                    dr("Pipe ID") = oCAEFceCitationInfoLocal.ID
                    dr("Deleted") = oCAEFceCitationInfoLocal.Deleted
                    dr("Created By") = oCAEFceCitationInfoLocal.CreatedBy
                    dr("Date Created") = oCAEFceCitationInfoLocal.CreatedOn
                    dr("Last Edited By") = oCAEFceCitationInfoLocal.ModifiedBy
                    dr("Date Last Edited") = oCAEFceCitationInfoLocal.ModifiedOn
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
        Private Sub CAEFceCitationInfoChanged(ByVal bolValue As Boolean) Handles oCAEFceCitationInfo.CAEFceCitationInfoChanged
            'RaiseEvent TemplateChanged(bolValue)
        End Sub
#End Region
    End Class
End Namespace
