' -------------------------------------------------------------------------------
' MUSTER.BusinessLogic.pFinancialActivity
' Provides the operations required to manipulate a FinancialActivity object.
' 
' Copyright (C) 2004, 2005 CIBER, Inc.
' All rights reserved.
' 
' Release   Initials    Date        Description
' 1.0         AB       06/23/2005    Original class definition
' 
' Function          Description
' -------------------------------------------------------------------------------
' Attribute          Description
' -------------------------------------------------------------------------------
Namespace MUSTER.BusinessLogic
    <Serializable()> Public Class pTecFinancialActivityPlanner

#Region "Private Member Variables"

        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private MusterException As Exceptions.MusterExceptions = New MUSTER.Exceptions.MusterExceptions
        Private colTecFinActivityPlanner As MUSTER.Info.TecFinancialActivityPlannerCollection = New MUSTER.Info.TecFinancialActivityPlannerCollection
        Private otecFinActivityPlannerInfo As MUSTER.Info.TecFinancialActivityPlannerInfo = New MUSTER.Info.TecFinancialActivityPlannerInfo
        Private oTecFinActivityPlannerDB As MUSTER.DataAccess.TecFinancialActivityPlannerDB = New MUSTER.DataAccess.TecFinancialActivityPlannerDB
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("Financial").ID
        Private nID As Int64 = -1
#End Region
#Region "Exposed Events"
        Public Delegate Sub TecFinActivityPlannerBLChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub TecFinActivityPlannerBLColChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub tecFinActivityPlannerBLErrEventHandler(ByVal MsgStr As String)
        Public Delegate Sub TecFinActivityPlannerInfoChanged()
        Public Event TecFinActivityPlannerBLChanged As TecFinActivityPlannerBLChangedEventHandler
        Public Event TecFinActivityPlannerBLColChanged As TecFinActivityPlannerBLColChangedEventHandler
        Public Event TecFinActivityPlannerBLErr As tecFinActivityPlannerBLErrEventHandler
        Public Event TecFinActivityPlannerInfChanged As TecFinActivityPlannerInfoChanged
        ' indicates change in the underlying _ProtoInfo structure
#End Region
#Region "Constructors"
        Public Sub New()
            otecFinActivityPlannerInfo = New MUSTER.Info.TecFinancialActivityPlannerInfo
        End Sub
        Public Sub New(ByVal TextID As Integer)
            otecFinActivityPlannerInfo = New MUSTER.Info.TecFinancialActivityPlannerInfo
            Me.Retrieve(TextID)
        End Sub

#End Region
#Region "Exposed Attributes"
        Public Property ID() As Int64
            Get
                Return otecFinActivityPlannerInfo.EventID
            End Get
            Set(ByVal Value As Int64)
                otecFinActivityPlannerInfo.EventID = Value
            End Set
        End Property

        Public Property ActivityTypeID() As Int64
            Get
                Return otecFinActivityPlannerInfo.ActivityTypeID
            End Get
            Set(ByVal Value As Int64)
                otecFinActivityPlannerInfo.ActivityTypeID = Value
            End Set
        End Property

        Public Property Cost() As Double
            Get
                Return otecFinActivityPlannerInfo.Cost
            End Get
            Set(ByVal Value As Double)
                otecFinActivityPlannerInfo.Cost = Value
            End Set
        End Property

        Public Property Duration() As Int64
            Get
                Return otecFinActivityPlannerInfo.Duration
            End Get
            Set(ByVal Value As Int64)
                otecFinActivityPlannerInfo.Duration = Value
            End Set
        End Property

        Public Property IsDirty() As Boolean
            Get
                Return otecFinActivityPlannerInfo.IsDirty
            End Get
            Set(ByVal Value As Boolean)
                otecFinActivityPlannerInfo.IsDirty = Value
            End Set
        End Property

#End Region
#Region "Exposed Operations"
#Region "General Operations"
        Public Sub Clear()
            otecFinActivityPlannerInfo = New MUSTER.Info.TecFinancialActivityPlannerInfo
        End Sub
        Public Sub Reset()
            otecFinActivityPlannerInfo.Reset()
        End Sub
#End Region
#Region "Info Operations"
        Public Function Retrieve(ByVal ID As Int64, Optional ByVal activityID As Int64 = -1) As MUSTER.Info.TecFinancialActivityPlannerInfo
            Try
                otecFinActivityPlannerInfo = oTecFinActivityPlannerDB.DBGetByID(ID, activityID)
                If otecFinActivityPlannerInfo.EventID = 0 Then
                    nID -= 1
                End If

                Return otecFinActivityPlannerInfo

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        ' Saves the data in the current Info object
        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal strModuleName As String = "")
            Try
                If Me.ValidateData(strModuleName) Then

                    oTecFinActivityPlannerDB.Put(otecFinActivityPlannerInfo, moduleID, staffID, returnVal)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If

                    otecFinActivityPlannerInfo.Archive()

                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        ' Validates the data before saving
        Public Function ValidateData(Optional ByVal [module] As String = "Financial") As Boolean
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
#End Region
#Region "Collection Operations"
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim xTecFinancialActivityPlannerInfo As MUSTER.Info.TecFinancialActivityPlannerInfo
            Try
                For Each xTecFinancialActivityPlannerInfo In Me.colTecFinActivityPlanner.Values
                    If xTecFinancialActivityPlannerInfo.IsDirty Then
                        otecFinActivityPlannerInfo = xTecFinancialActivityPlannerInfo
                        Me.Save(moduleID, staffID, returnVal)
                    End If
                Next
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        ' Removes the entity supplied from the collection
        Public Sub Add(ByRef oFinActivityInfoInfo As MUSTER.Info.TecFinancialActivityPlannerInfo)
            Try
                colTecFinActivityPlanner.Add(oFinActivityInfoInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        ' Removes the entity supplied from the collection
        Public Sub Remove(ByVal oTecFinActivityPlannerInfo As Object)
            Try
                colTecFinActivityPlanner.Remove(oTecFinActivityPlannerInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        ' Gets all the info
        Public Function GetAll(ByVal nReason As Int64) As MUSTER.Info.TecFinancialActivityPlannerCollection
            Try
                colTecFinActivityPlanner.Clear()
                colTecFinActivityPlanner = Me.oTecFinActivityPlannerDB.DBGetAll
                Return colTecFinActivityPlanner
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function


#End Region
#Region " Populate Routines "


        Public Function PopulateTecActivityListForCosts(Optional ByVal eventID As Int64 = -1) As DataTable
            Dim dtReturn As DataTable
            Try

                Dim criteria As String = String.Empty

                If eventID <> -1 Then
                    criteria = String.Format(" Where Property_ID not in (Select [Activity_Type_id] as Property_ID from tblTECFIN_Event_Activity_Planner Where [Tec_Event_ID] = {0})", _
                                               eventID)
                End If

                dtReturn = GetDataTable(String.Format("(Select Property_ID, Property_Name, 1 as PROPERTY_POSITION from VTECActivityList A   LEFT JOIN tblTEC_Activity_Plus P on A.Property_ID = P.Activity_ID  Where  isnull(CostMode,0) <> 0) L {0} ", criteria))

                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function PopulateTechnicalActivityList(Optional ByVal eventID As Int64 = -1) As DataTable
            Try

                Dim criteria As String = String.Empty

                If eventID <> -1 Then
                    criteria = String.Format(" Where Property_ID not in (Select [Activity_Type_id] as Property_ID from tblTECFIN_Event_Activity_Planner Where [Tec_Event_ID] = {0})", _
                                               eventID)
                End If

                Dim dtReturn As DataTable = GetDataTable(String.Format("VTECActivityList {0}", criteria), False, False)

                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function PopulateActivityPlanning(ByVal eventID As Int64) As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable(String.Format("vTECFINActivityPlanningList where ([Event ID] = {0} or {0} = 0) ", eventID), False, False)
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function PopulateCostFormats() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vFinancialCostFormats", True, True)
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function


        Public Function GetDataTable(ByVal strProperty As String, Optional ByVal PropertyMaster As Boolean = True, Optional ByVal IncludeBlank As Boolean = False) As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String

            strSQL = ""
            If PropertyMaster Then
                If IncludeBlank Then
                    strSQL = " SELECT '' as PROPERTY_NAME, 0 as PROPERTY_ID, 0 as PROPERTY_POSITION "
                    strSQL &= " UNION "
                End If

                strSQL &= " SELECT PROPERTY_NAME, PROPERTY_ID, PROPERTY_POSITION FROM " & strProperty
                strSQL &= " order by 1 "
            Else
                strSQL &= " SELECT * FROM " & strProperty
            End If
            Try
                dsReturn = oTecFinActivityPlannerDB.DBGetDS(strSQL)
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
#Region "Private Operations"
        Public Function GetNext() As String
            Return GetNextPrev(1)
        End Function
        Public Function GetPrevious() As String
            Return GetNextPrev(-1)
        End Function
        Private Function GetNextPrev(ByVal direction As Integer) As String
            Dim strArr() As String = colTecFinActivityPlanner.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ActivityTypeID))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colTecFinActivityPlanner.Item(nArr.GetValue(colIndex + direction)).ActivityTypeID.ToString
            Else
                Return colTecFinActivityPlanner.Item(nArr.GetValue(colIndex)).ActivityTypeID.ToString
            End If
        End Function
#End Region

    End Class
End Namespace


