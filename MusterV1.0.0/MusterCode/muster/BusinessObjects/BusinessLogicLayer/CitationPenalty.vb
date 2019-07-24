'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.CitationPenalty
'   Provides the operations required to manipulate an Entity object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0      RAF/MKK     06/27/2005  Original class definition
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
' NOTE: This file to be used as CitationPenalty to build other objects.
'       Replace keyword "CitationPenalty" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pCitationPenalty
#Region "Public Events"
        Public Event CitationPenaltyErr(ByVal MsgStr As String)
        Public Event CitationPenaltyChanged(ByVal bolValue As Boolean)
        Public Event ColChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        Private WithEvents oCitationPenaltyInfo As Muster.Info.CitationPenaltyInfo
        Private WithEvents colCitationPenaltys As Muster.Info.CitationPenaltysCollection
        Private oCitationPenaltyDB As New Muster.DataAccess.CitationPenaltyDB
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private nID As Int64 = -1
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
#End Region
#Region "Constructors"
        Public Sub New()
            oCitationPenaltyInfo = New Muster.Info.CitationPenaltyInfo
            colCitationPenaltys = New Muster.Info.CitationPenaltysCollection
        End Sub
        '********************************************************
        '
        ' Overloaded NEW which will populate with a single instance
        '   of the named CitationPenalty object.
        '
        '********************************************************
        Public Sub New(ByVal CitationPenaltyName As String)
            oCitationPenaltyInfo = New Muster.Info.CitationPenaltyInfo
            colCitationPenaltys = New Muster.Info.CitationPenaltysCollection
            Me.Retrieve(CitationPenaltyName)
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Integer
            Get
                Return oCitationPenaltyInfo.ID
            End Get
            Set(ByVal Value As Integer)
                oCitationPenaltyInfo.ID = Integer.Parse(Value)
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oCitationPenaltyInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oCitationPenaltyInfo.Deleted = Boolean.Parse(Value)
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oCitationPenaltyInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oCitationPenaltyInfo.IsDirty = Boolean.Parse(value)
            End Set
        End Property
        Public Property colIsDirty() As Boolean
            Get
                Dim xCitationPenaltyinfo As MUSTER.Info.CitationPenaltyInfo
                For Each xCitationPenaltyinfo In colCitationPenaltys.Values
                    If xCitationPenaltyinfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
            Set(ByVal Value As Boolean)
                oCitationPenaltyInfo.IsDirty = Value
            End Set
        End Property
        Public Property StateCitation() As String
            Get
                Return oCitationPenaltyInfo.StateCitation
            End Get
            Set(ByVal Value As String)
                oCitationPenaltyInfo.StateCitation = Value
            End Set
        End Property
        Public Property Small() As Integer
            Get
                Return oCitationPenaltyInfo.Small
            End Get
            Set(ByVal Value As Integer)
                oCitationPenaltyInfo.Small = Value
            End Set
        End Property
        Public Property Section() As String
            Get
                Return oCitationPenaltyInfo.Section
            End Get
            Set(ByVal Value As String)
                oCitationPenaltyInfo.Section = Value
            End Set
        End Property
        Public Property Medium() As Integer
            Get
                Return oCitationPenaltyInfo.Medium
            End Get
            Set(ByVal Value As Integer)
                oCitationPenaltyInfo.Medium = Value
            End Set
        End Property
        Public Property Large() As Integer
            Get
                Return oCitationPenaltyInfo.Large
            End Get
            Set(ByVal Value As Integer)
                oCitationPenaltyInfo.Large = Value
            End Set
        End Property
        Public Property FederalCitation() As String
            Get
                Return oCitationPenaltyInfo.FederalCitation
            End Get
            Set(ByVal Value As String)
                oCitationPenaltyInfo.FederalCitation = Value
            End Set
        End Property
        Public Property EPA() As String
            Get
                Return oCitationPenaltyInfo.EPA
            End Get
            Set(ByVal Value As String)
                oCitationPenaltyInfo.EPA = Value
            End Set
        End Property
        Public Property Description() As String
            Get
                Return oCitationPenaltyInfo.Description
            End Get
            Set(ByVal Value As String)
                oCitationPenaltyInfo.Description = Value
            End Set
        End Property
        Public Property DATE_LAST_EDITED() As Date
            Get
                Return oCitationPenaltyInfo.DATE_LAST_EDITED
            End Get
            Set(ByVal Value As Date)
                oCitationPenaltyInfo.DATE_LAST_EDITED = Value
            End Set
        End Property
        Public Property LAST_EDITED_BY() As String
            Get
                Return oCitationPenaltyInfo.LAST_EDITED_BY
            End Get
            Set(ByVal Value As String)
                oCitationPenaltyInfo.LAST_EDITED_BY = Value
            End Set
        End Property
        Public Property DATE_CREATED() As Date
            Get
                Return oCitationPenaltyInfo.DATE_CREATED
            End Get
            Set(ByVal Value As Date)
                oCitationPenaltyInfo.DATE_CREATED = Value
            End Set
        End Property
        Public Property CREATED_BY() As String
            Get
                Return oCitationPenaltyInfo.CREATED_BY
            End Get
            Set(ByVal Value As String)
                oCitationPenaltyInfo.CREATED_BY = Value
            End Set
        End Property
        Public Property CorrectiveAction() As String
            Get
                Return oCitationPenaltyInfo.CorrectiveAction
            End Get
            Set(ByVal Value As String)
                oCitationPenaltyInfo.CorrectiveAction = Value
            End Set
        End Property
        Public Property Category() As String
            Get
                Return oCitationPenaltyInfo.Category
            End Get
            Set(ByVal Value As String)
                oCitationPenaltyInfo.Category = Value
            End Set
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        'Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal ID As Integer) As Muster.Info.CitationPenaltyInfo
            Try
                oCitationPenaltyInfo = colCitationPenaltys.Item(ID)
                If Not oCitationPenaltyInfo Is Nothing Then
                    Return oCitationPenaltyInfo
                End If
                oCitationPenaltyInfo = oCitationPenaltyDB.DBGetByID(ID)
                If oCitationPenaltyInfo.ID = 0 Then
                    oCitationPenaltyInfo.ID = nID
                    nID -= 1
                End If
                colCitationPenaltys.Add(oCitationPenaltyInfo)
                Return oCitationPenaltyInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Saves the data in the current Info object
        Public Function Save(Optional ByVal bolValidated As Boolean = False) As Boolean
            Try
                If Not bolValidated Then
                    If Not Me.ValidateData() Then
                        Return False
                    End If
                End If
                If Not (oCitationPenaltyInfo.ID < 0 And oCitationPenaltyInfo.Deleted) Then
                    Dim OldKey As String = oCitationPenaltyInfo.ID.ToString
                    oCitationPenaltyDB.Put(oCitationPenaltyInfo)
                    If Not bolValidated Then
                        If oCitationPenaltyInfo.ID.ToString <> OldKey Then
                            colCitationPenaltys.ChangeKey(OldKey, oCitationPenaltyInfo.ID.ToString)
                        End If
                    End If
                    oCitationPenaltyInfo.Archive()
                    oCitationPenaltyInfo.IsDirty = False
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Validates the data before saving
        Public Function ValidateData(Optional ByVal [module] As String = "Registration") As Boolean
            Dim errStr As String = ""
            Dim validateSuccess As Boolean = False
            Try

                If errStr.Length > 0 Or Not validateSuccess Then
                    RaiseEvent CitationPenaltyErr(errStr)
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
        Function GetAll() As MUSTER.Info.CitationPenaltysCollection
            Try
                colCitationPenaltys.Clear()
                colCitationPenaltys = oCitationPenaltyDB.GetAllInfo
                Return colCitationPenaltys
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal ID As Integer)
            Try
                oCitationPenaltyInfo = oCitationPenaltyDB.DBGetByID(ID)
                If oCitationPenaltyInfo.ID = 0 Then
                    oCitationPenaltyInfo.ID = nID
                    nID -= 1
                End If
                colCitationPenaltys.Add(oCitationPenaltyInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Function Add(ByRef CitationPenaltyInfo As MUSTER.Info.CitationPenaltyInfo) As Boolean
            Try
                oCitationPenaltyInfo = CitationPenaltyInfo
                If ValidateData() Then
                    If oCitationPenaltyInfo.ID <= 0 Then
                        oCitationPenaltyInfo.ID = nID
                        nID -= 1
                    End If
                    colCitationPenaltys.Add(oCitationPenaltyInfo)
                    Return True
                Else
                    Return False
                End If
                colCitationPenaltys.Add(oCitationPenaltyInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Integer)
            Dim myIndex As Int16 = 1
            Dim oCitationPenaltyInfoLocal As MUSTER.Info.CitationPenaltyInfo

            Try
                For Each oCitationPenaltyInfoLocal In colCitationPenaltys.Values
                    If oCitationPenaltyInfoLocal.ID = ID Then
                        colCitationPenaltys.Remove(oCitationPenaltyInfoLocal)
                        Exit Sub
                    End If
                    myIndex += 1
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("CitationPenalty " & ID.ToString & " is not in the collection of CitationPenaltys.")
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oCitationPenalty As MUSTER.Info.CitationPenaltyInfo)
            Try
                colCitationPenaltys.Remove(oCitationPenalty)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("CitationPenalty " & oCitationPenalty.ID & " is not in the collection of CitationPenaltys.")
        End Sub
        Public Sub Flush()
            Dim xCitationPenaltyInfo As MUSTER.Info.CitationPenaltyInfo
            For Each xCitationPenaltyInfo In colCitationPenaltys.Values
                If xCitationPenaltyInfo.IsDirty Then
                    oCitationPenaltyInfo = xCitationPenaltyInfo
                    Me.Save(True)
                End If
            Next
        End Sub
#End Region
#Region "General Operations"
        Public Function Clear()
            oCitationPenaltyInfo = New MUSTER.Info.CitationPenaltyInfo
        End Function
        Public Function Reset()
            oCitationPenaltyInfo.Reset()
        End Function
#End Region
#Region "Miscellaneous Operations"
        '7/20 - Replace the whole Function
        Public Function EntityTable() As DataTable
            Dim oCitationPenaltyInfoLocal As New MUSTER.Info.CitationPenaltyInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Try

                tbEntityTable.Columns.Add("Citation_ID", GetType(Integer))
                tbEntityTable.Columns.Add("StateCitation", GetType(String))
                tbEntityTable.Columns.Add("FederalCitation", GetType(String))
                tbEntityTable.Columns.Add("Section", GetType(String))
                tbEntityTable.Columns.Add("Category", GetType(String))
                tbEntityTable.Columns.Add("CitationText", GetType(String))
                tbEntityTable.Columns.Add("Small", GetType(Integer))
                tbEntityTable.Columns.Add("Medium", GetType(Integer))
                tbEntityTable.Columns.Add("Large", GetType(Integer))
                tbEntityTable.Columns.Add("CorrectiveAction", GetType(String))
                tbEntityTable.Columns.Add("EPA", GetType(String))
                tbEntityTable.Columns.Add("CreatedBy", GetType(String))
                tbEntityTable.Columns.Add("CreatedOn", GetType(Date))
                tbEntityTable.Columns.Add("ModifiedBy", GetType(String))
                tbEntityTable.Columns.Add("ModifiedOn", GetType(Date))
                tbEntityTable.Columns.Add("Deleted", GetType(Boolean))


                For Each oCitationPenaltyInfoLocal In colCitationPenaltys.Values
                    dr = tbEntityTable.NewRow()
                    dr("Citation_ID") = oCitationPenaltyInfoLocal.ID
                    dr("StateCitation") = oCitationPenaltyInfoLocal.StateCitation
                    dr("FederalCitation") = oCitationPenaltyInfoLocal.FederalCitation
                    dr("Section") = oCitationPenaltyInfoLocal.Section
                    dr("Category") = oCitationPenaltyInfoLocal.Category
                    dr("CitationText") = oCitationPenaltyInfoLocal.Description
                    dr("Small") = oCitationPenaltyInfoLocal.Small
                    dr("Medium") = oCitationPenaltyInfoLocal.Medium
                    dr("Large") = oCitationPenaltyInfoLocal.Large
                    dr("CorrectiveAction") = oCitationPenaltyInfoLocal.CorrectiveAction
                    dr("EPA") = oCitationPenaltyInfoLocal.EPA
                    dr("CreatedBy") = oCitationPenaltyInfoLocal.CREATED_BY
                    dr("CreatedOn") = oCitationPenaltyInfoLocal.DATE_CREATED
                    dr("ModifiedBy") = oCitationPenaltyInfoLocal.LAST_EDITED_BY
                    dr("ModifiedOn") = oCitationPenaltyInfoLocal.DATE_LAST_EDITED
                    dr("Deleted") = oCitationPenaltyInfoLocal.Deleted
                    tbEntityTable.Rows.Add(dr)
                Next
                Return tbEntityTable
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region
#End Region
#Region "External Event Handlers"
        Private Sub CitationPenaltyInfoChanged(ByVal bolValue As Boolean) Handles oCitationPenaltyInfo.CitationPenaltyInfoChanged
            RaiseEvent CitationPenaltyChanged(bolValue)
        End Sub
        Private Sub CitationPenaltyColChanged(ByVal bolValue As Boolean) Handles colCitationPenaltys.CitationPenaltyColChanged
            RaiseEvent ColChanged(bolValue)
        End Sub
#End Region
    End Class
End Namespace
