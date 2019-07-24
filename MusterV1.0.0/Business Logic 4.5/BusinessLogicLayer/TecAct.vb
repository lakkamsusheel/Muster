Imports MUSTER.Info.TecActInfo

Namespace MUSTER.BusinessLogic
    <Serializable()> Public Class pTecAct
        '-------------------------------------------------------------------------------
        ' MUSTER.BusinessLogic.TecAct
        '   Provides the operations required to manipulate a Technical Activity object.
        '
        ' Copyright (C) 2004, 2005 CIBER, Inc.
        ' All rights reserved.
        '
        ' Release   Initials    Date        Description
        '  1.0         JC       5/31/2005    Original class definition
        '  1.1         JC       7/28/2005    Added InUse to detect association of activity
        '                                       to existing LUST events.
        ' Function          Description
        '-------------------------------------------------------------------------------
        ' Attribute          Description
        '-------------------------------------------------------------------------------
#Region "Private Member Variables"
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private colTecActs As MUSTER.Info.TecActCollection = New MUSTER.Info.TecActCollection
        Private MusterException As Exceptions.MusterExceptions = New MUSTER.Exceptions.MusterExceptions
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("Tech Act").ID
        Private nID As Int64 = -1
        Private oTecActDB As MUSTER.DataAccess.TecActDB = New MUSTER.DataAccess.TecActDB
        Private oTecActInfo As MUSTER.Info.TecActInfo = New MUSTER.Info.TecActInfo
#End Region
#Region "Public Events"
        Public Delegate Sub TecActBLColChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub TecActChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub TecActErrEventHandler(ByVal MsgStr As String)
        ' indicates change in the underlying TecActInfo structure
        Public Delegate Sub TecActInfoChanged()

        Public Event TecActChanged As TecActChangedEventHandler
        Public Event TecActColChanged As TecActBLColChangedEventHandler
        Public Event TecActErr As TecActErrEventHandler
#End Region
#Region "Constructors"
        Public Sub New()
            oTecActInfo = New MUSTER.Info.TecActInfo
        End Sub
        Public Sub New(ByVal ActivityID As Integer)
            oTecActInfo = New MUSTER.Info.TecActInfo

            Me.Retrieve(ActivityID)
        End Sub
        Public Sub New(ByVal ActivityName As String)
            oTecActInfo = New MUSTER.Info.TecActInfo
            Me.Retrieve(ActivityName)
        End Sub
#End Region
#Region "Exposed Attributes"
        ' Gets/Sets the active flag for the technical document (from TecDoc.Active)
        Public Property Active() As Boolean
            Get
                Return oTecActInfo.Active()
            End Get
            Set(ByVal Value As Boolean)
                oTecActInfo.Active = Value
            End Set
        End Property

        Public Property Cost() As Double
            Get
                Return oTecActInfo.Cost
            End Get
            Set(ByVal Value As Double)
                oTecActInfo.Cost = Value
            End Set
        End Property

        Public Property colIsDirty() As Boolean
            Get
                Dim xTecActInfo As MUSTER.Info.TecActInfo
                For Each xTecActInfo In colTecActs.Values
                    If xTecActInfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
            Set(ByVal Value As Boolean)
                oTecActInfo.IsDirty = Value
            End Set
        End Property


        Public Property CostMode() As ActivityCostModeEnum
            Get
                Return oTecActInfo.CostMode
            End Get
            Set(ByVal Value As ActivityCostModeEnum)
                oTecActInfo.CostMode = Value
            End Set
        End Property


        ' The ID of the user that created the row
        Public Property CreatedBy() As String
            Get
                Return oTecActInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oTecActInfo.CreatedBy = Value
            End Set
        End Property
        ' The date on which the row was created
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oTecActInfo.CreatedOn
            End Get
        End Property
        ' Indicates the deleted state of the row
        Public Property Deleted() As Boolean
            Get
                Return oTecActInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oTecActInfo.Deleted = Value
            End Set
        End Property
        ' The entity ID associated with a technical document.
        Public ReadOnly Property EntityID() As Integer
            Get
                Return oTecActInfo.EntityID
            End Get
        End Property
        ' Gets/Sets the warn days threshold
        Public Property WarnDays() As String
            Get
                Return oTecActInfo.WarnDays
            End Get
            Set(ByVal Value As String)
                oTecActInfo.WarnDays = Value
            End Set
        End Property
        ' Gets the action days threshold
        Public Property ActDays() As Long
            Get
                Return oTecActInfo.ActDays
            End Get
            Set(ByVal Value As Long)
                oTecActInfo.ActDays = Value
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oTecActInfo.IsDirty()
            End Get
            Set(ByVal Value As Boolean)
                oTecActInfo.IsDirty = Value
            End Set
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oTecActInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oTecActInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oTecActInfo.ModifiedOn
            End Get
        End Property
        ' Gets/Sets the name of the technical document (from oTecDoc.Name)
        Public Property Name() As String
            Get
                Return oTecActInfo.Name
            End Get
            Set(ByVal Value As String)
                oTecActInfo.Name = Value
            End Set
        End Property

        ' Gets/Sets the name of the technical document (from oTecDoc.Name)
        Public Property Activity_ID() As Int64
            Get
                Return oTecActInfo.ID
            End Get
            Set(ByVal Value As Int64)
                oTecActInfo.ID = Value
            End Set
        End Property

        ' Gets/Sets the name of the technical document (from oTecDoc.Name)
        Public Property DocumentsCollection() As MUSTER.Info.TecDocCollection
            Get
                Return oTecActInfo.DocumentsCollection
            End Get
            Set(ByVal Value As MUSTER.Info.TecDocCollection)
                oTecActInfo.DocumentsCollection = Value
            End Set
        End Property

        ' Gets/Sets the name of the technical document (from oTecDoc.Name)
        Public ReadOnly Property InfoObject() As MUSTER.Info.TecActInfo
            Get
                Return oTecActInfo
            End Get
        End Property
#End Region
#Region "Exposed Methods"
#Region "General Operations"
        Public Sub Clear()
            oTecActInfo = New MUSTER.Info.TecActInfo
        End Sub
        Public Sub Reset()
            oTecActInfo.Reset()
        End Sub
#End Region
#Region "Info Operations"
        ' Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal ActID As Int64) As MUSTER.Info.TecActInfo
            Try
                oTecActInfo = oTecActDB.DBGetByID(ActID)
                If oTecActInfo.ID = 0 Then
                    nID -= 1
                End If

                Return oTecActInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        ' Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal ActName As String) As MUSTER.Info.TecActInfo

            Try
                oTecActInfo = oTecActDB.DBGetByName(ActName)
                If oTecActInfo.ID = 0 Then
                    nID -= 1
                End If

                Return oTecActInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        ' Saves the data in the current Info object
        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim strModuleName As String = String.Empty
            Try
                If Me.ValidateData(strModuleName) Then
                    oTecActDB.Put(oTecActInfo, moduleID, staffID, returnVal)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                    oTecActInfo.Archive()

                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Validates the data before saving
        Public Function ValidateData(Optional ByVal [module] As String = "Registration") As Boolean
            Dim errStr As String = ""
            Dim validateSuccess As Boolean = True

            Try
                Select Case [module]
                    Case "Registration"

                        ' if any validations failed
                        Exit Select
                    Case "Technical"

                        ' if any validations failed
                        Exit Select
                End Select
                'If errStr.Length > 0 Or Not validateSuccess Then
                '    RaiseEvent LustEventErr(errStr)
                'End If
                Return validateSuccess
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function InUse() As Boolean
            Return oTecActDB.InUse(oTecActInfo)
        End Function
#End Region
#Region "Collection Operations"

        ' Removes the entity supplied from the collection
        Public Sub Add(ByRef oTecActInfo As MUSTER.Info.TecActInfo)
            Try
                colTecActs.Add(oTecActInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub

        ' Gets all the info
        Public Function GetAll() As MUSTER.Info.TecActCollection
            Try
                colTecActs.Clear()
                colTecActs = oTecActDB.GetAllInfo
                Return colTecActs
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function

        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim xTecActInfo As MUSTER.Info.TecActInfo
            Try
                For Each xTecActInfo In colTecActs.Values
                    If xTecActInfo.IsDirty Then
                        oTecActInfo = xTecActInfo
                        Me.Save(moduleID, staffID, returnVal)
                    End If
                Next
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try

        End Sub

        ' Removes the entity supplied from the collection
        Public Sub Remove(ByVal oTecActInfo)
            Try
                colTecActs.Remove(oTecActInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Function GetNext() As String
            Return GetNextPrev(1)
        End Function
        Public Function GetPrevious() As String
            Return GetNextPrev(-1)
        End Function
        Private Function GetNextPrev(ByVal direction As Integer) As String
            Dim strArr() As String = colTecActs.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.Activity_ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colTecActs.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            Else
                Return colTecActs.Item(nArr.GetValue(colIndex)).ID.ToString
            End If
        End Function
#End Region
#Region " Populate Routines "

        Public Function PopulateTecActivityListForCosts() As DataTable
            Dim dtReturn As DataTable
            Try
                dtReturn = GetDataTable("(Select Property_ID, Property_Name, 1 as PROPERTY_POSITION from VTECActivityList  A   LEFT JOIN tblTEC_Activity_Plus P on A.Property_ID = P.Activity_ID  Where  isnull(CostMode,0) <> 0) L ")

                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function PopulateTecActivityList(Optional ByVal IncludeBlank As Boolean = False, Optional ByVal ShowAll As Boolean = True) As DataTable
            Dim dtReturn As DataTable
            Try
                If ShowAll Then
                    dtReturn = GetDataTable("VTECActivityList", IncludeBlank)
                Else
                    dtReturn = GetDataTable("VTECActivityList_Active", IncludeBlank)
                End If

                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function PopulateCostModes() As DataTable
            Dim dtReturn As DataTable
            Try
                dtReturn = GetDataTable("vTecActCostModes")

                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        'Public Function PopulateTecDocumentTypes() As DataTable
        '    Try
        '        Dim dtReturn As DataTable = GetDataTable("VTECDocumentTypes")
        '        Return dtReturn
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Function
        Public Function GetDataTable(ByVal strProperty As String, Optional ByVal IncludeBlank As Boolean = False) As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String

            strSQL = ""
            If IncludeBlank Then
                strSQL = " SELECT '' as PROPERTY_NAME, 0 as PROPERTY_ID "
                strSQL &= " UNION "
            End If

            strSQL &= " SELECT PROPERTY_NAME, PROPERTY_ID FROM " & strProperty
            strSQL &= " order by 1 "
            Try
                dsReturn = oTecActDB.DBGetDS(strSQL)
                If dsReturn.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsReturn.Tables(0)
                End If
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
#End Region
#End Region
    End Class
End Namespace
