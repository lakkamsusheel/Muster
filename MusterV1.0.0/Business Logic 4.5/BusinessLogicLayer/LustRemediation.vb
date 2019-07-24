'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.LustRemediation
'   Provides the operations required to manipulate a Remediation System object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0         JC       4/12/2005    Original class definition
'
' Function          Description
'-------------------------------------------------------------------------------
' Attribute          Description
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    ' the remediation system associated to the LUST event.  There must be a remediation activity associated to the LUST event in order to associate a remediation system.
    <Serializable()> _
    Public Class pLustRemediation

#Region "Public Events"

#End Region

#Region "Private Member Variables"
        Private WithEvents oLustRemediationInfo As New MUSTER.Info.LustRemediationInfo
        Private WithEvents colLustRemediation As New MUSTER.Info.LustRemediationCollection
        Private oLustRemediationDB As New MUSTER.DataAccess.LustRemediationDB
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private nID As Int64 = -1
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("Pipe").ID

#End Region

#Region "Constructors"
        Public Sub New()
            oLustRemediationInfo = New MUSTER.Info.LustRemediationInfo
            ColLustRemediation = New MUSTER.Info.LustRemediationCollection
        End Sub

#End Region

#Region "Exposed Attributes"
        ' The remediation system ID for the current Info
        Public Property ID() As Long
            Get
                Return oLustRemediationInfo.ID
            End Get
            Set(ByVal Value As Long)
                oLustRemediationInfo.ID = Value
            End Set
        End Property

        Public Property SystemSequence() As Long
            Get
                Return oLustRemediationInfo.SystemSequence
            End Get
            Set(ByVal Value As Long)
                oLustRemediationInfo.SystemSequence = Value
            End Set
        End Property

        Public Property SystemDeclaration() As Long
            Get
                Return oLustRemediationInfo.SystemDeclaration
            End Get
            Set(ByVal Value As Long)
                oLustRemediationInfo.SystemDeclaration = Value
            End Set
        End Property
        ' The date the remediation system was placed in use (from info.DatePlacedInUse).
        Public Property DateInUse() As Date
            Get
                Return oLustRemediationInfo.DatePlacedInUse
            End Get
            Set(ByVal Value As Date)
                oLustRemediationInfo.DatePlacedInUse = Value
            End Set
        End Property
        ' The type of remediation system (from info.RemSysType)
        Public Property RemSysType() As Long
            Get
                Return oLustRemediationInfo.RemSysType
            End Get
            Set(ByVal Value As Long)
                oLustRemediationInfo.RemSysType = Value
            End Set
        End Property
        ' The description of the remediation system (from info.Description)
        Public Property Description() As String
            Get
                Return oLustRemediationInfo.Description
            End Get
            Set(ByVal Value As String)
                oLustRemediationInfo.Description = Value
            End Set
        End Property
        Public Property Manufacturer() As String
            Get
                Return oLustRemediationInfo.Manufacturer
            End Get
            Set(ByVal Value As String)
                oLustRemediationInfo.Manufacturer = Value
            End Set
        End Property
        Public Property OWSAgeofComp() As String
            Get
                Return oLustRemediationInfo.OWSAgeofComp
            End Get
            Set(ByVal Value As String)
                oLustRemediationInfo.OWSAgeofComp = Value
            End Set
        End Property
        Public Property OWSManName() As String
            Get
                Return oLustRemediationInfo.OWSManName
            End Get
            Set(ByVal Value As String)
                oLustRemediationInfo.OWSManName = Value
            End Set
        End Property
        Public Property OWSModelNumber() As String
            Get
                Return oLustRemediationInfo.OWSModelNumber
            End Get
            Set(ByVal Value As String)
                oLustRemediationInfo.OWSModelNumber = Value
            End Set
        End Property
        Public Property OWSNewUsed() As Integer
            Get
                Return oLustRemediationInfo.OWSNewUsed
            End Get
            Set(ByVal Value As Integer)
                oLustRemediationInfo.OWSNewUsed = Value
            End Set
        End Property
        Public Property OWSSerialNumber() As String
            Get
                Return oLustRemediationInfo.OWSSerialNumber
            End Get
            Set(ByVal Value As String)
                oLustRemediationInfo.OWSSerialNumber = Value
            End Set
        End Property
        Public Property OWSSize() As String
            Get
                Return oLustRemediationInfo.OWSSize
            End Get
            Set(ByVal Value As String)
                oLustRemediationInfo.OWSSize = Value
            End Set
        End Property
        Public Property MotorAgeofComp() As String
            Get
                Return oLustRemediationInfo.MotorAgeofComp
            End Get
            Set(ByVal Value As String)
                oLustRemediationInfo.MotorAgeofComp = Value
            End Set
        End Property


        Public Property MotorManName() As String
            Get
                Return oLustRemediationInfo.MotorManName
            End Get
            Set(ByVal Value As String)
                oLustRemediationInfo.MotorManName = Value
            End Set
        End Property
        Public Property MotorModelNumber() As String
            Get
                Return oLustRemediationInfo.MotorModelNumber
            End Get
            Set(ByVal Value As String)
                oLustRemediationInfo.MotorModelNumber = Value
            End Set
        End Property
        Public Property MotorNewUsed() As Integer
            Get
                Return oLustRemediationInfo.MotorNewUsed
            End Get
            Set(ByVal Value As Integer)
                oLustRemediationInfo.MotorNewUsed = Value
            End Set
        End Property
        Public Property MotorSerialNumber() As String
            Get
                Return oLustRemediationInfo.MotorSerialNumber
            End Get
            Set(ByVal Value As String)
                oLustRemediationInfo.MotorSerialNumber = Value
            End Set
        End Property
        Public Property MotorSize() As String
            Get
                Return oLustRemediationInfo.MotorSize
            End Get
            Set(ByVal Value As String)
                oLustRemediationInfo.MotorSize = Value
            End Set
        End Property

        Public Property StripperAgeofComp() As String
            Get
                Return oLustRemediationInfo.StripperAgeofComp
            End Get
            Set(ByVal Value As String)
                oLustRemediationInfo.StripperAgeofComp = Value
            End Set
        End Property
        Public Property StripperManName() As String
            Get
                Return oLustRemediationInfo.StripperManName
            End Get
            Set(ByVal Value As String)
                oLustRemediationInfo.StripperManName = Value
            End Set
        End Property
        Public Property StripperModelNumber() As String
            Get
                Return oLustRemediationInfo.StripperModelNumber
            End Get
            Set(ByVal Value As String)
                oLustRemediationInfo.StripperModelNumber = Value
            End Set
        End Property
        Public Property StripperNewUsed() As Integer
            Get
                Return oLustRemediationInfo.StripperNewUsed
            End Get
            Set(ByVal Value As Integer)
                oLustRemediationInfo.StripperNewUsed = Value
            End Set
        End Property
        Public Property StripperSerialNumber() As String
            Get
                Return oLustRemediationInfo.StripperSerialNumber
            End Get
            Set(ByVal Value As String)
                oLustRemediationInfo.StripperSerialNumber = Value
            End Set
        End Property
        Public Property StripperSize() As String
            Get
                Return oLustRemediationInfo.StripperSize
            End Get
            Set(ByVal Value As String)
                oLustRemediationInfo.StripperSize = Value
            End Set
        End Property

        Public Property VacPump1AgeofComp() As String
            Get
                Return oLustRemediationInfo.VacPump1AgeofComp
            End Get
            Set(ByVal Value As String)
                oLustRemediationInfo.VacPump1AgeofComp = Value
            End Set
        End Property
        Public Property VacPump1ManName() As String
            Get
                Return oLustRemediationInfo.VacPump1ManName
            End Get
            Set(ByVal Value As String)
                oLustRemediationInfo.VacPump1ManName = Value
            End Set
        End Property
        Public Property VacPump1ModelNumber() As String
            Get
                Return oLustRemediationInfo.VacPump1ModelNumber
            End Get
            Set(ByVal Value As String)
                oLustRemediationInfo.VacPump1ModelNumber = Value
            End Set
        End Property
        Public Property VacPump1NewUsed() As Integer
            Get
                Return oLustRemediationInfo.VacPump1NewUsed
            End Get
            Set(ByVal Value As Integer)
                oLustRemediationInfo.VacPump1NewUsed = Value
            End Set
        End Property
        Public Property VacPump1SerialNumber() As String
            Get
                Return oLustRemediationInfo.VacPump1SerialNumber
            End Get
            Set(ByVal Value As String)
                oLustRemediationInfo.VacPump1SerialNumber = Value
            End Set
        End Property
        Public Property VacPump1Size() As String
            Get
                Return oLustRemediationInfo.VacPump1Size
            End Get
            Set(ByVal Value As String)
                oLustRemediationInfo.VacPump1Size = Value
            End Set
        End Property
        Public Property VacPump1Seal() As Integer
            Get
                Return oLustRemediationInfo.VacPump1Seal
            End Get
            Set(ByVal Value As Integer)
                oLustRemediationInfo.VacPump1Seal = Value
            End Set
        End Property

        Public Property VacPump2AgeofComp() As String
            Get
                Return oLustRemediationInfo.VacPump2AgeofComp
            End Get
            Set(ByVal Value As String)
                oLustRemediationInfo.VacPump2AgeofComp = Value
            End Set
        End Property
        Public Property VacPump2ManName() As String
            Get
                Return oLustRemediationInfo.VacPump2ManName
            End Get
            Set(ByVal Value As String)
                oLustRemediationInfo.VacPump2ManName = Value
            End Set
        End Property
        Public Property VacPump2ModelNumber() As String
            Get
                Return oLustRemediationInfo.VacPump2ModelNumber
            End Get
            Set(ByVal Value As String)
                oLustRemediationInfo.VacPump2ModelNumber = Value
            End Set
        End Property
        Public Property VacPump2NewUsed() As Integer
            Get
                Return oLustRemediationInfo.VacPump2NewUsed
            End Get
            Set(ByVal Value As Integer)
                oLustRemediationInfo.VacPump2NewUsed = Value
            End Set
        End Property
        Public Property VacPump2SerialNumber() As String
            Get
                Return oLustRemediationInfo.VacPump2SerialNumber
            End Get
            Set(ByVal Value As String)
                oLustRemediationInfo.VacPump2SerialNumber = Value
            End Set
        End Property
        Public Property VacPump2Size() As String
            Get
                Return oLustRemediationInfo.VacPump2Size
            End Get
            Set(ByVal Value As String)
                oLustRemediationInfo.VacPump2Size = Value
            End Set
        End Property
        Public Property VacPump2Seal() As Integer
            Get
                Return oLustRemediationInfo.VacPump2Seal
            End Get
            Set(ByVal Value As Integer)
                oLustRemediationInfo.VacPump2Seal = Value
            End Set
        End Property
        ' The boolean indicating if the remediation system is owned (true) or leased (false) (from info.Owned).
        Public Property Owned() As Integer
            Get
                Return oLustRemediationInfo.Owned
            End Get
            Set(ByVal Value As Integer)
                oLustRemediationInfo.Owned = Value
            End Set
        End Property
        Public Property Owner() As String
            Get
                Return oLustRemediationInfo.Owner
            End Get
            Set(ByVal Value As String)
                oLustRemediationInfo.Owner = Value
            End Set
        End Property
        ' The buildingsize for the remediation system (from info.BuildingSize)
        Public Property BuildingSize() As String
            Get
                Return oLustRemediationInfo.BuildingSize
            End Get
            Set(ByVal Value As String)
                oLustRemediationInfo.BuildingSize = Value
            End Set
        End Property
        ' The boolean indicating if the system is trailer mounted (true) or skid mounted (false) (from info.TrailerMount).
        Public Property MountType() As Integer
            Get
                Return oLustRemediationInfo.MountType
            End Get
            Set(ByVal Value As Integer)
                oLustRemediationInfo.MountType = Value
            End Set
        End Property
        ' The date the remediation system was last refurbished (from info.RefurbDate)
        Public Property RefurbDate() As Date
            Get
                Return oLustRemediationInfo.RefurbDate
            End Get
            Set(ByVal Value As Date)
                oLustRemediationInfo.RefurbDate = Value
            End Set
        End Property
        ' Additional notes that are associated with the remediation system (from info.Notes)
        Public Property Notes() As String
            Get
                Return oLustRemediationInfo.Notes
            End Get
            Set(ByVal Value As String)
                oLustRemediationInfo.Notes = Value
            End Set
        End Property
        ' The date the remediation system was purchased (if Owned) (from info.PurchaseDate)
        Public Property PurchaseDate() As Date
            Get
                Return oLustRemediationInfo.PurchaseDate
            End Get
            Set(ByVal Value As Date)
                oLustRemediationInfo.PurchaseDate = Value
            End Set
        End Property
        Public Property CreatedBy() As String
            Get
                Return oLustRemediationInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oLustRemediationInfo.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oLustRemediationInfo.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oLustRemediationInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oLustRemediationInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oLustRemediationInfo.ModifiedOn
            End Get
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oLustRemediationInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oLustRemediationInfo.Deleted = Value
            End Set
        End Property
        ' The first option for the remediation system (from info.Option1)
        Public Property Option1() As Long
            Get
                Return oLustRemediationInfo.Option1
            End Get
            Set(ByVal Value As Long)
                oLustRemediationInfo.Option1 = Value
            End Set
        End Property
        ' The second option for the remediation system (from info.Option2)
        Public Property Option2() As Long
            Get
                Return oLustRemediationInfo.Option2
            End Get
            Set(ByVal Value As Long)
                oLustRemediationInfo.Option2 = Value
            End Set
        End Property
        ' The third option for the remediation system (from info.Option3)
        Public Property Option3() As Long
            Get
                Return oLustRemediationInfo.Option3
            End Get
            Set(ByVal Value As Long)
                oLustRemediationInfo.Option3 = Value
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oLustRemediationInfo.IsDirty
            End Get
            Set(ByVal Value As Boolean)
                oLustRemediationInfo.IsDirty = Value
            End Set
        End Property

        ' The remediation system info for the currently indexed remediation system in the collection
        Public Property oRemSys() As Info.LustRemediationInfo
            Get
                Return oRemSys
            End Get
            Set(ByVal Value As Info.LustRemediationInfo)
                oRemSys = Value
            End Set
        End Property
        ' The remediation system DB interface object
        Private Property oRemSysDB() As DataAccess.LustRemediationDB
            Get
                Return oRemSysDB
            End Get
            Set(ByVal Value As DataAccess.LustRemediationDB)
                oRemSysDB = Value
            End Set
        End Property
        ' The collection of remediation system infos
        Private Property RemSysCol() As Info.LustRemediationCollection
            Get
                Return colLustRemediation
            End Get
            Set(ByVal Value As Info.LustRemediationCollection)
                colLustRemediation = Value
            End Set
        End Property
#End Region

#Region "Exposed Operations"
#Region "Info Operations"
        Public Function Retrieve(ByVal id As Int64, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.LustRemediationInfo
            Dim oLustRemediationInfoLocal As MUSTER.Info.LustRemediationInfo
            Try
                For Each oLustRemediationInfoLocal In colLustRemediation.Values
                    If oLustRemediationInfoLocal.ID = id Then
                        oLustRemediationInfo = oLustRemediationInfoLocal
                        Return oLustRemediationInfo
                    End If
                Next
                oLustRemediationInfo = oLustRemediationDB.DBGetByID(id)
                If oLustRemediationInfo.ID = 0 Then
                    nID -= 1
                End If
                colLustRemediation.Add(oLustRemediationInfo)
                Return oLustRemediationInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal bolValidated As Boolean = False)
            Dim strModuleName As String = String.Empty
            Try
                If Me.ValidateData(strModuleName) Then
                    oLustRemediationDB.Put(oLustRemediationInfo, moduleID, staffID, returnVal)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                    oLustRemediationInfo.Archive()
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub

        ' Validates the data before saving
        Public Function ValidateData(Optional ByVal [module] As String = "Technical") As Boolean

            Return True

        End Function
#End Region
#Region "Collection Operations"

        Public Function GetAll(ByVal LustEvent As Int64) As MUSTER.Info.LustRemediationCollection
            Try
                colLustRemediation.Clear()
                colLustRemediation = oLustRemediationDB.DBGetByEventID(LustEvent)
                Return colLustRemediation
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        ' Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef id As Int64)
            Try
                oLustRemediationInfo = oLustRemediationDB.DBGetByID(id)
                If oLustRemediationInfo.ID = 0 Then
                    nID -= 1
                End If
                colLustRemediation.Add(oLustRemediationInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        ' Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oRemediationSystem As MUSTER.Info.LustRemediationInfo)
            Try
                oLustRemediationInfo = oRemediationSystem
                colLustRemediation.Add(oLustRemediationInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        ' Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Long)
            Dim oLustRemediationInfo As MUSTER.Info.LustRemediationInfo
            Try
                oLustRemediationInfo = colLustRemediation.Item(ID)
                If Not (oLustRemediationInfo Is Nothing) Then
                    colLustRemediation.Remove(oLustRemediationInfo)
                    Exit Sub
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Function RemoveAll() As Object
            Try
                colLustRemediation.Clear()
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim xLustRemediationInfo As MUSTER.Info.LustRemediationInfo
            For Each xLustRemediationInfo In colLustRemediation.Values
                If xLustRemediationInfo.IsDirty Then
                    oLustRemediationInfo = xLustRemediationInfo
                    Me.Save(moduleID, staffID, returnVal)
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
            Dim strArr() As String = colLustRemediation.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colLustRemediation.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            Else
                Return colLustRemediation.Item(nArr.GetValue(colIndex)).ID.ToString
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oLustRemediationInfo = New MUSTER.Info.LustRemediationInfo
        End Sub
        Public Sub Reset()
            oLustRemediationInfo.Reset()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function CheckRemediationSystemPermissions(ByVal nRemSystemId As Integer) As Boolean
            Dim dsReturn As New DataSet
            Dim strSQL As String
            Try
                strSQL = "select * from vREMEDIATION_PERMISSIONS where rem_system_id=" & nRemSystemId.ToString
                dsReturn = oLustRemediationDB.DBGetDS(strSQL)
                If dsReturn.Tables(0).Rows.Count > 0 Then
                    Return False
                Else
                    Return True
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
#End Region

#End Region

#Region "Lookup/Dataset Operations"
        Public Function HistoricalSystemsDataset() As DataSet
            Dim dsRemSys As New DataSet
            Dim drRow As DataRow
            Dim oCol As DataColumn
            Dim dsRel As DataRelation
            Dim strSQL As String
            Try

                'strSQL = "SELECT * FROM V_LUST_TANK_DISPLAY_DATA WHERE FACILITY_ID = '" & oFacilityInfo.ID.ToString & "' ORDER BY [TANK SITE ID];" & _
                '        "SELECT * FROM V_LUST_PIPE_DISPLAY_DATA WHERE FACILITY_ID = '" & oFacilityInfo.ID.ToString & "' ORDER BY [PIPE SITE ID] "


                strSQL = "select * from vLUSTREMEDIATIONCURRENT ;"
                strSQL = strSQL & " "
                strSQL = strSQL & "select * from vLUSTREMEDIATIONHISTORICAL ;"


                dsRemSys = oLustRemediationDB.DBGetDS(strSQL)

                dsRemSys.Tables(0).DefaultView.Sort = "Start_Date DESC"
                dsRemSys.Tables(1).DefaultView.Sort = "Start_Date DESC"

                dsRel = New DataRelation("CurrentToHistory", dsRemSys.Tables(0).Columns("SYSTEM_SEQ"), dsRemSys.Tables(1).Columns("SYSTEM_SEQ"), False)

                dsRemSys.Relations.Add(dsRel)
                Return dsRemSys
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function AvailableSystemsDataset() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vLUSTREMEDIATIONAVAILABLESYSTEMS")
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function


        Public Function PopulateRemediationType() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vLUSTREMEDIATIONTYPE")
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function PopulateRemediationOptionalEquipment() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vLUSTREMEDIATIONOPTIONALEQ")
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function PopulateRemediationOwnedLeased() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vLUSTREMEDIATIONOWNEDLEASED")
                Return dtReturn

            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function PopulateRemediationNewUsed() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vLUSTREMEDIATIONNEWUSED")
                Return dtReturn

            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function PopulateRemediationMountType() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vLUSTREMEDIATIONMOUNTTYPE")
                Return dtReturn

            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function PopulateRemediationPumpSeal() As DataTable
            Try

                Dim dtReturn As DataTable = GetDataTable("vLUSTREMEDIATIONPUMPTYPE")
                Return dtReturn

            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function SystemPreviousLocations(ByVal nSystemSeq As Int64) As DataTable
            Dim dsRemSys As New DataSet
            Dim dtRemSys As New DataTable
            Dim drRow As DataRow
            Dim oCol As DataColumn
            Dim dsRel As DataRelation
            Dim strSQL As String
            Try

                'strSQL = "SELECT * FROM V_LUST_TANK_DISPLAY_DATA WHERE FACILITY_ID = '" & oFacilityInfo.ID.ToString & "' ORDER BY [TANK SITE ID];" & _
                '        "SELECT * FROM V_LUST_PIPE_DISPLAY_DATA WHERE FACILITY_ID = '" & oFacilityInfo.ID.ToString & "' ORDER BY [PIPE SITE ID] "


                strSQL = "select distinct Facility_ID, cast(Facility_ID as varchar) + ' - ' +  Facility_Name as Full_Name from vLUSTREMEDIATIONHISTORICAL where Facility_ID is not null and System_SEQ = " & nSystemSeq

                dsRemSys = oLustRemediationDB.DBGetDS(strSQL)

                If dsRemSys.Tables(0).Rows.Count > 0 Then
                    dtRemSys = dsRemSys.Tables(0)
                End If

                Return dtRemSys
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function


        Public Function GetDataTable(ByVal strProperty As String, Optional ByVal nVal As Int64 = 0, Optional ByVal bolDistinct As Boolean = False) As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String
            If bolDistinct Then
                strSQL = "SELECT DISTINCT PROPERTY_ID, PROPERTY_NAME FROM " + strProperty
            Else
                strSQL = "SELECT * FROM " & strProperty
            End If
            If nVal <> 0 Then
                strSQL = strSQL + " WHERE PROPERTY_ID_PARENT = " + nVal.ToString()
            End If
            Try
                dsReturn = oLustRemediationDB.DBGetDS(strSQL)
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


    End Class
End Namespace
