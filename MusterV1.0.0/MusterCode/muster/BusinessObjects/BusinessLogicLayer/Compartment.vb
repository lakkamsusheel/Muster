'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.Compartment
'   Provides the info and collection objects to the client for manipulating
'   an AddressInfo object.
'   Copyright (C) 2004 CIBER, Inc.
'  All rights reserved.
'
' Release   Initials    Date        Description
'   1.0         EN      12/13/04    Original class definition.
'   1.1         EN      12/24/04    Added properties to the collection in set method of each properties.
'   1.2         AN      01/03/05    Added Try catch and Exception Handling/Logging
'   1.3         EN      01/06/05    Modified Reset Method. 
'   1.4         MNR     01/25/05    Added GetNext(), GetPrevious() and GetNextPrev(..) functions
'   1.5         AN      02/02/05    Added Comments object
'   1.6         EN      02/10/05    Added checkCAPStatusPipe 
'   1.7         AN      02/11/05    Made Change on Retrieve to fix problem with passing object as nothing
'   1.8         AB      02/17/05    Added DataAge check to the Retrieve function
'   1.9         MNR     03/02/05    Modified Flush to save compartments in the order they were entered in the Grid
'                                   compartment with compNum -1 will save before compartment with compNum -2
'   2.0         MNR     03/15/05    Added Load Sub
'   2.1         MNR     03/16/05    Removed strSrc from events
'
' Function          Description
' Retrieve(FullKey,ShowDeleted)   Returns the compartmentinfo object based on the key..
' GetAll()    Returns an  CompartmentCollection with all Person objects
' Add(ID)           Adds the Compartment identified by arg ID to the 
'                           internal  CompartmentCollection
' Add(Name)         Adds the Compartment identified by arg NAME to the internal 
'                    CompartmentCollection            
' Add(Compartment)       Adds the Compartment passed as the argument to the internal 
'                            CompartmentCollection
' Remove(ID)        Removes the Compartment identified by arg ID from the internal 
'                            CompartmentCollection
' Remove(NAME)      Removes the Compartment identified by arg NAME from the 
'                           internal  CompartmentCollection
' CompartmentTable()     Returns a datatable containing all columns for the Person 

' colIsDirty()       Returns a boolean indicating whether any of the Compartmentinfo
'                    objects in the CompartmentCollection has been modified since the
'                    last time it was retrieved from/saved to the repository.
' Flush()            Marshalls all modified/added Compartmentinfo objects in the 
'                        CompartmentCollection to the repository.
' Save()             Marshalls the internal Compartmentinfo object to the repository.

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pCompartment
#Region "Public Events"
        Public Event evtFacCapStatus(ByVal facID As Integer)
        Public Event evtCompartmentErr(ByVal MsgStr As String)
        Public Event evtCompartmentChanged(ByVal bolValue As Boolean)
        Public Event evtCompartmentsChanged(ByVal bolValue As Boolean)
        'Public Event evtPipeCommentsChanged(ByVal bolValue As Boolean)
        'Added by Elango 
        'Public Event evtCAPStatusPipe(ByVal nVal As Integer)
        'Added by kiran
        'Public Event evtCompColTank(ByVal TankID As Integer, ByVal compartmentCol As MUSTER.Info.CompartmentCollection)
        'Public Event evtPipeColTank(ByVal TankID As Integer, ByVal pipeCol As MUSTER.Info.PipesCollection)
        'Public Event evtPipeCommentsCol(ByVal pipeID As Integer, ByVal compID As Integer, ByVal tankID As Integer, ByVal commentsCol As MUSTER.Info.CommentsCollection)
        'end changes
        'Public Event evtCompInfoTank(ByVal compartmentInfo As MUSTER.Info.CompartmentInfo, ByVal strDesc As String)
        'Public Event evtPipeInfoCompartment(ByVal pipeInfo As MUSTER.Info.PipeInfo, ByVal strDesc As String)
        'Public Event evtTankInfoCompCol(ByRef colComp As MUSTER.Info.CompartmentCollection)
        'Public Event evtCompInfoCompID(ByVal cmpID As String)
        'Public Event evtCompartmentChangeKey(ByVal oldID As String, ByVal newID As String)
        'Public Event evtSyncPipeInCol(ByVal pipeInfo As MUSTER.Info.PipeInfo)
#End Region
#Region "Private Member Variables"
        Private oTankInfo As MUSTER.Info.TankInfo
        'Private colCompartment As MUSTER.Info.CompartmentCollection
        Private WithEvents oCompartmentInfo As MUSTER.Info.CompartmentInfo
        Private oCompartmentDB As New MUSTER.DataAccess.CompartmentDB
        Private WithEvents oCompartmentPipe As MUSTER.BusinessLogic.pPipe
        Private WithEvents oProperty As MUSTER.BusinessLogic.pProperty
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private strNewID As String
        Private nID As Int64 = -1
        Private bolShowDeleted As Boolean
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Public Sub New(Optional ByRef TankInfo As MUSTER.Info.TankInfo = Nothing)
            If TankInfo Is Nothing Then
                oTankInfo = New MUSTER.Info.TankInfo
            Else
                oTankInfo = TankInfo
            End If
            oCompartmentInfo = New MUSTER.Info.CompartmentInfo
            'colCompartment = New MUSTER.Info.CompartmentCollection
            oCompartmentPipe = New MUSTER.BusinessLogic.pPipe(oTankInfo)
            oProperty = New MUSTER.BusinessLogic.pProperty
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As String
            Get
                Return oCompartmentInfo.ID
            End Get
            Set(ByVal value As String)
                oCompartmentInfo.ID = value
            End Set
        End Property
        Public Property TankId() As Integer
            Get
                Return oCompartmentInfo.TankId
            End Get
            Set(ByVal value As Integer)
                Dim oldTankId As Integer = oCompartmentInfo.TankId
                oCompartmentInfo.TankId = value
                Dim oPipeLocal As MUSTER.Info.PipeInfo
                For Each oPipeLocal In oCompartmentPipe.PipeCollection.Values
                    If oPipeLocal.TankID = oldTankId And oPipeLocal.CompartmentNumber = oCompartmentInfo.COMPARTMENTNumber Then
                        If oPipeLocal.TankID <= 0 Then
                            oPipeLocal.TankID = oCompartmentInfo.TankId
                        End If
                    End If
                Next
            End Set
        End Property
        Public Property COMPARTMENTNumber() As Integer
            Get
                Return oCompartmentInfo.COMPARTMENTNumber
            End Get
            Set(ByVal value As Integer)
                Dim oldCompNum As Integer = oCompartmentInfo.COMPARTMENTNumber
                oCompartmentInfo.COMPARTMENTNumber = value
                Dim oPipeLocal As MUSTER.Info.PipeInfo
                For Each oPipeLocal In oCompartmentPipe.PipeCollection.Values
                    If oPipeLocal.TankID = oCompartmentInfo.TankId And oPipeLocal.CompartmentNumber = oldCompNum Then
                        oPipeLocal.CompartmentNumber = oCompartmentInfo.COMPARTMENTNumber
                    End If
                Next
            End Set
        End Property
        Public Property Capacity() As Integer
            Get
                Return oCompartmentInfo.Capacity
            End Get
            Set(ByVal value As Integer)
                oCompartmentInfo.Capacity = value
            End Set
        End Property
        Public Property CCERCLA() As Integer
            Get
                Return oCompartmentInfo.CCERCLA
            End Get
            Set(ByVal value As Integer)
                oCompartmentInfo.CCERCLA = value
                Dim oPipeLocal As MUSTER.Info.PipeInfo
                For Each oPipeLocal In oTankInfo.pipesCollection.Values
                    If oPipeLocal.CompartmentID = oCompartmentInfo.ID Then
                        oPipeLocal.CompartmentCERCLA = oCompartmentInfo.CCERCLA
                    End If
                Next
            End Set
        End Property
        Public Property Substance() As Integer
            Get
                Return oCompartmentInfo.Substance
            End Get
            Set(ByVal value As Integer)
                oCompartmentInfo.Substance = value
                Dim oPipeLocal As MUSTER.Info.PipeInfo
                For Each oPipeLocal In oTankInfo.pipesCollection.Values
                    If oPipeLocal.CompartmentID = oCompartmentInfo.ID Then
                        oPipeLocal.CompartmentSubstance = oCompartmentInfo.Substance
                    End If
                Next
            End Set
        End Property
        Public Property FuelTypeId() As Integer
            Get
                Return oCompartmentInfo.FuelTypeId
            End Get
            Set(ByVal value As Integer)
                oCompartmentInfo.FuelTypeId = value
                Dim oPipeLocal As MUSTER.Info.PipeInfo
                For Each oPipeLocal In oTankInfo.pipesCollection.Values
                    If oPipeLocal.CompartmentID = oCompartmentInfo.ID Then
                        oPipeLocal.CompartmentFuelType = oCompartmentInfo.FuelTypeId
                    End If
                Next
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oCompartmentInfo.Deleted
            End Get

            Set(ByVal value As Boolean)
                oCompartmentInfo.Deleted = value
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oCompartmentInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oCompartmentInfo.IsDirty = value
            End Set
        End Property
        Public Property CreatedBy() As String
            Get
                Return oCompartmentInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oCompartmentInfo.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oCompartmentInfo.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oCompartmentInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oCompartmentInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oCompartmentInfo.ModifiedOn
            End Get
        End Property
        Public Property colIsDirty() As Boolean
            Get
                Dim xCompartmentInfo As MUSTER.Info.CompartmentInfo
                For Each xCompartmentInfo In oTankInfo.CompartmentCollection.Values
                    If xCompartmentInfo.IsDirty Then
                        'MsgBox("BLL F:" + oTankInfo.FacilityId.ToString + " TI:" + oTankInfo.TankIndex.ToString + " CID:" + xCompartmentInfo.ID)
                        Return True
                        Exit Property
                    End If
                Next
                If oCompartmentPipe.colIsDirty Then
                    Return True
                    Exit Property
                End If
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
        Public Property FacilityId() As Integer
            Get
                Return oCompartmentInfo.FacilityId
            End Get
            Set(ByVal Value As Integer)
                oCompartmentInfo.FacilityId = Value
                Dim oPipeLocal As MUSTER.Info.PipeInfo
                For Each oPipeLocal In oCompartmentPipe.PipeCollection.Values
                    If oPipeLocal.CompartmentID = oCompartmentInfo.ID Then
                        If oPipeLocal.FacilityID = 0 Then
                            oPipeLocal.FacilityID = oCompartmentInfo.FacilityId
                        End If
                    End If
                Next
            End Set
        End Property
        Public Property TankSiteID() As Integer
            Get
                Return oCompartmentInfo.TankSiteID
            End Get
            Set(ByVal Value As Integer)
                oCompartmentInfo.TankSiteID = Value
                Dim oPipeLocal As MUSTER.Info.PipeInfo
                For Each oPipeLocal In oCompartmentPipe.PipeCollection.Values
                    If oPipeLocal.CompartmentID = oCompartmentInfo.ID Then
                        oPipeLocal.TankSiteID = oCompartmentInfo.TankSiteID
                    End If
                Next
            End Set
        End Property
        Public ReadOnly Property FuelTypeIdDesc() As String
            Get
                Return oProperty.Retrieve(Me.FuelTypeId).Name
            End Get
        End Property
        Public ReadOnly Property SubstanceDesc() As String
            Get
                Return oProperty.Retrieve(Me.Substance).Name
            End Get
        End Property
        Public ReadOnly Property CERCLADesc() As String
            Get
                ' TODO - have to put CERCLA values in property_master table
                'Return oProperty.Retrieve(Me.CCERCLA).Name
                Dim dsLocal As DataSet
                Dim strSQL As String = "SELECT * FROM vCERLATYPE WHERE CASRN = '" + CCERCLA.ToString + "'"
                dsLocal = oCompartmentDB.DBGetDS(strSQL)
                If dsLocal.Tables(0).Rows.Count > 0 Then
                    Return dsLocal.Tables(0).Rows(0).Item("Substance").ToString
                Else
                    Return String.Empty
                End If
            End Get
        End Property
        Public Property TankInfo() As MUSTER.Info.TankInfo
            Get
                Return oTankInfo
            End Get
            Set(ByVal Value As MUSTER.Info.TankInfo)
                oTankInfo = Value
            End Set
        End Property
        Public ReadOnly Property CompInfo() As MUSTER.Info.CompartmentInfo
            Get
                Return oCompartmentInfo
            End Get
        End Property
        ' Collections
        Public ReadOnly Property CompartmentCollection() As MUSTER.Info.CompartmentCollection
            Get
                Return oTankInfo.CompartmentCollection
            End Get
        End Property
        Public ReadOnly Property PipeCollection() As MUSTER.Info.PipesCollection
            Get
                Return oCompartmentPipe.PipeCollection
            End Get
        End Property
        Public Property Pipes() As MUSTER.BusinessLogic.pPipe
            Get
                Return oCompartmentPipe
            End Get
            Set(ByVal Value As MUSTER.BusinessLogic.pPipe)
                oCompartmentPipe = Value
            End Set
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        Public Sub Load(ByRef TankInfo As MUSTER.Info.TankInfo, ByRef ds As DataSet, ByVal [Module] As String)
            Dim dr As DataRow
            Dim hasComp As Boolean = False
            oTankInfo = TankInfo
            Try
                If ds.Tables("Compartments").Rows.Count > 0 Then
                    For Each dr In ds.Tables("Compartments").Select("TANK_ID = " + oTankInfo.TankId.ToString)
                        oCompartmentInfo = New MUSTER.Info.CompartmentInfo(dr)
                        oCompartmentInfo.TankSiteID = oTankInfo.TankIndex
                        oCompartmentInfo.FacilityId = oTankInfo.FacilityId
                        'RaiseEvent evtCompInfoTank(oCompartmentInfo, "ADD")
                        oTankInfo.CompartmentCollection.Add(oCompartmentInfo)
                        ds.Tables("Compartments").Rows.Remove(dr)
                        hasComp = True
                    Next
                End If
                If Not hasComp Then
                    Add(New MUSTER.Info.CompartmentInfo)
                End If
                oCompartmentPipe.Load(oTankInfo, ds, [Module])
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function Retrieve(ByRef TankInfo As MUSTER.Info.TankInfo, ByVal FullKey As String, Optional ByVal ShowDeleted As Boolean = False) As MUSTER.Info.CompartmentInfo
            Dim bolDataAged As Boolean = False
            Dim strArray() As String
            oTankInfo = TankInfo
            Try
                'Dim colCompartmentContained As MUSTER.Info.CompartmentCollection
                'RaiseEvent evtTankInfoCompCol(colCompartmentContained)
                If oTankInfo.CompartmentCollection.Contains(FullKey) Then
                    oCompartmentInfo = oTankInfo.CompartmentCollection.Item(FullKey)
                    If oCompartmentInfo.IsDirty = False And oCompartmentInfo.IsAgedData = True Then
                        bolDataAged = True
                    Else
                        Return oCompartmentInfo
                    End If
                End If
                If bolDataAged Then
                    oTankInfo.CompartmentCollection.Remove(oCompartmentInfo)
                    'RaiseEvent evtCompInfoTank(oCompartmentInfo, "REMOVE")
                End If
                strArray = FullKey.Split("|")
                oCompartmentInfo = oCompartmentDB.DBGetByKey(strArray, ShowDeleted)
                Add(oCompartmentInfo)
                'oTankInfo.CompartmentCollection.Add(oCompartmentInfo)
                'RaiseEvent evtCompInfoTank(oCompartmentInfo, "ADD")
                For Each opipeInfo As MUSTER.Info.PipeInfo In oTankInfo.pipesCollection.Values
                    If opipeInfo.CompartmentID = oCompartmentInfo.ID Then
                        If opipeInfo.SubstanceDesc <> oCompartmentInfo.Substance Then
                            opipeInfo.SubstanceDesc = oCompartmentInfo.Substance
                        End If
                        opipeInfo.CompartmentCERCLA = oCompartmentInfo.CCERCLA
                        opipeInfo.CompartmentFuelType = oCompartmentInfo.FuelTypeId
                    End If
                Next
                Return oCompartmentInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function Retrieve(ByRef TankInfo As MUSTER.Info.TankInfo, ByVal tnkID As Integer, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.CompartmentInfo
            Dim oCompartmentInfoLocal As MUSTER.Info.CompartmentInfo
            Dim bolDataAged As Boolean = False
            Dim oPipeInfoLocal As MUSTER.Info.PipeInfo
            Try
                oTankInfo = TankInfo
                Dim BolRetrievePipe As Boolean = False
                ' check in collection
                If oTankInfo.CompartmentCollection.Count > 0 Then
                    For Each oCompartmentInfoLocal In oTankInfo.CompartmentCollection.Values
                        If oCompartmentInfoLocal.TankId = tnkID Then
                            oCompartmentInfo = oCompartmentInfoLocal
                            oCompartmentPipe.Retrieve(oTankInfo, oTankInfo.TankId, oCompartmentInfo, showDeleted)
                            For Each opipeInfo As MUSTER.Info.PipeInfo In oTankInfo.pipesCollection.Values
                                If opipeInfo.CompartmentID = oCompartmentInfo.ID Then
                                    If opipeInfo.SubstanceDesc <> oCompartmentInfo.Substance Then
                                        opipeInfo.SubstanceDesc = oCompartmentInfo.Substance
                                    End If
                                    opipeInfo.CompartmentCERCLA = oCompartmentInfo.CCERCLA
                                    opipeInfo.CompartmentFuelType = oCompartmentInfo.FuelTypeId
                                End If
                            Next
                            If oCompartmentInfo.IsDirty = False And oCompartmentInfo.IsAgedData = True Then
                                bolDataAged = True
                                BolRetrievePipe = False
                            Else
                                BolRetrievePipe = True
                            End If
                        End If
                    Next
                End If
                If bolDataAged Then
                    oTankInfo.CompartmentCollection.Remove(oCompartmentInfo)
                End If
                'Get From DB
                If Not BolRetrievePipe Then
                    oTankInfo.CompartmentCollection = oCompartmentDB.DBGetByTankID(tnkID, showDeleted)
                    If oTankInfo.CompartmentCollection.Count > 0 Then
                        For Each oCompartmentInfoLocal In oTankInfo.CompartmentCollection.Values
                            oCompartmentInfo = oCompartmentInfoLocal
                            oCompartmentInfo.TankSiteID = oTankInfo.TankIndex
                            oCompartmentInfo.FacilityId = oTankInfo.FacilityId
                            oCompartmentPipe.Retrieve(oTankInfo, oCompartmentInfo.TankId, oCompartmentInfo, showDeleted)
                            For Each opipeInfo As MUSTER.Info.PipeInfo In oTankInfo.pipesCollection.Values
                                If opipeInfo.CompartmentID = oCompartmentInfo.ID Then
                                    If opipeInfo.SubstanceDesc <> oCompartmentInfo.Substance Then
                                        opipeInfo.SubstanceDesc = oCompartmentInfo.Substance
                                    End If
                                    opipeInfo.CompartmentCERCLA = oCompartmentInfo.CCERCLA
                                    opipeInfo.CompartmentFuelType = oCompartmentInfo.FuelTypeId
                                End If
                            Next
                        Next
                    Else
                        'if not in DB, create new instance
                        Add(New MUSTER.Info.CompartmentInfo)
                        'oCompartmentInfo.TankId = tnkID
                        'oCompartmentInfo.TankSiteID = oTankInfo.TankIndex
                        'oTankInfo.CompartmentCollection.Add(oCompartmentInfo)
                    End If
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Return oCompartmentInfo
        End Function
        Public Function Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal strUser As String, Optional ByVal bolValidated As Boolean = False, Optional ByVal bolDelete As Boolean = False, Optional ByVal bolSaveAsInspection As Boolean = False) As Boolean
            Dim oldID As String
            Try
                If Not (oCompartmentInfo.COMPARTMENTNumber < 0 And oCompartmentInfo.Deleted) Then
                    oldID = oCompartmentInfo.ID
                    oCompartmentDB.Put(oCompartmentInfo, moduleID, staffID, returnVal, strUser)
                    If Not returnVal = String.Empty Then
                        Exit Function
                    End If
                    If Not bolValidated Then
                        If CType(oldID.Split("|")(1), Integer) <> oCompartmentInfo.COMPARTMENTNumber Then
                            oTankInfo.CompartmentCollection.ChangeKey(oldID, oCompartmentInfo.ID)
                            For Each oPipeLocal As MUSTER.Info.PipeInfo In TankInfo.pipesCollection.Values
                                If oPipeLocal.CompartmentID = oldID Then
                                    oPipeLocal.CompartmentID = oCompartmentInfo.ID
                                End If
                            Next
                        End If
                    End If
                    oCompartmentInfo.Archive()
                    oCompartmentInfo.IsDirty = False
                    Dim userID As String = String.Empty
                    If oCompartmentInfo.COMPARTMENTNumber <= 0 Then
                        userID = oCompartmentInfo.CreatedBy
                    Else
                        userID = oCompartmentInfo.ModifiedBy
                    End If
                    oCompartmentPipe.Flush(moduleID, staffID, returnVal, userID, bolSaveAsInspection)
                    If Not returnVal = String.Empty Then
                        Exit Function
                    End If
                End If
                If Not bolValidated And bolDelete Then
                    If oCompartmentInfo.Deleted Then
                        Dim strNext As String = Me.GetNext()
                        Dim strPrev As String = Me.GetPrevious()
                        If strNext = oCompartmentInfo.ID Then
                            If strPrev = oCompartmentInfo.ID Then
                                RaiseEvent evtCompartmentErr("Compartment: " + oCompartmentInfo.COMPARTMENTNumber.ToString + " of Tank: " + oCompartmentInfo.TankSiteID.ToString + " deleted")
                                oTankInfo.CompartmentCollection.Remove(oCompartmentInfo)
                                'RaiseEvent evtCompInfoTank(oCompartmentInfo, "REMOVE")
                                If bolDelete Then
                                    oCompartmentInfo = New MUSTER.Info.CompartmentInfo
                                Else
                                    oCompartmentInfo = Me.Retrieve(oTankInfo, oCompartmentInfo.TankId)
                                End If
                            Else
                                RaiseEvent evtCompartmentErr("Compartment: " + oCompartmentInfo.COMPARTMENTNumber.ToString + " of Tank: " + oCompartmentInfo.TankSiteID.ToString + " deleted")
                                oTankInfo.CompartmentCollection.Remove(oCompartmentInfo)
                                'RaiseEvent evtCompInfoTank(oCompartmentInfo, "REMOVE")
                                oCompartmentInfo = Me.Retrieve(oTankInfo, strPrev)
                            End If
                        Else
                            RaiseEvent evtCompartmentErr("Compartment: " + oCompartmentInfo.COMPARTMENTNumber.ToString + " of Tank: " + oCompartmentInfo.TankSiteID.ToString + " deleted")
                            oTankInfo.CompartmentCollection.Remove(oCompartmentInfo)
                            'RaiseEvent evtCompInfoTank(oCompartmentInfo, "REMOVE")
                            oCompartmentInfo = Me.Retrieve(oTankInfo, strNext)
                        End If
                    End If
                End If
                RaiseEvent evtCompartmentChanged(Me.IsDirty)
                Return True
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
                Return False
            End Try
        End Function
        '''Public Function DeleteCompartment() As Boolean
        '''    Try
        '''        For Each opipeinfo As MUSTER.Info.PipeInfo In oTankInfo.pipesCollection.Values
        '''            If opipeinfo.CompartmentID = oCompartmentInfo.ID Then
        '''                RaiseEvent evtCompartmentErr("The Specified compartment has associated Pipe(s). Delete Pipe(s) before deleting the compartment")
        '''                Exit Try
        '''            End If
        '''        Next
        '''        ' comp does not have pipe(s), delete comp
        '''        oCompartmentInfo.Deleted = True
        '''        Return False 'Me.Save(True, True)
        '''    Catch ex As Exception
        '''        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '''        Throw ex
        '''        Return False
        '''    End Try
        '''End Function
#End Region
#Region "Collection Operations"
        Public Function GetAll(Optional ByVal TankId As Integer = 0, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.CompartmentCollection
            Try
                oTankInfo.CompartmentCollection.Clear()
                oTankInfo.CompartmentCollection = oCompartmentDB.DBGetAllInfo(TankId, showDeleted)
                'RaiseEvent evtCompColTank(TankId, oTankInfo.CompartmentCollection)
                Return oTankInfo.CompartmentCollection
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an address to the collection as supplied by the caller
        Public Sub Add(ByRef oCompartment As MUSTER.Info.CompartmentInfo)
            Try
                oCompartmentInfo = oCompartment
                If oCompartmentInfo.ID = "0|0" Then
                    oCompartmentInfo.ID = oTankInfo.TankId & "|" & nID
                    nID -= 1
                End If
                oCompartmentInfo.TankSiteID = oTankInfo.TankIndex
                oCompartmentInfo.FacilityId = oTankInfo.FacilityId
                oTankInfo.CompartmentCollection.Add(oCompartmentInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oCompartmentInf As MUSTER.Info.CompartmentInfo)
            Try
                oTankInfo.CompartmentCollection.Remove(oCompartmentInf)
                'RaiseEvent evtCompInfoTank(oCompartmentInf, "REMOVE")
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

            Throw New Exception("Compartment Info " & oCompartmentInf.ID & " is not in the collection of Compartment Collection.")
        End Sub
        'Removes the Compartment called for by ID from the collection
        Public Sub Remove(ByVal FullKey As String)
            Dim myIndex As Int16 = 1
            Dim oCompartmentInfoLocal As MUSTER.Info.CompartmentInfo
            Try
                'Dim colCompartmentContained As MUSTER.Info.CompartmentCollection
                'RaiseEvent evtTankInfoCompCol(colCompartmentContained)
                For Each oCompartmentInfoLocal In oTankInfo.CompartmentCollection.Values
                    If oCompartmentInfoLocal.ID = FullKey Then
                        oTankInfo.CompartmentCollection.Remove(oCompartmentInfoLocal)
                        'RaiseEvent evtCompInfoTank(oCompartmentInfoLocal, "REMOVE")
                        Exit Sub
                    End If
                    myIndex += 1
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("Compartment " & ID.ToString & " is not in the collection of Persona.")
        End Sub
        '''Public Function colIsDirty() As Boolean
        '''    Dim oTempInfo As Muster.Info.CompartmentInfo
        '''    For Each oTempInfo In colCompartment.Values
        '''        If oTempInfo.IsDirty Then
        '''            Return True
        '''        End If
        '''    Next
        '''    Return False

        '''End Function
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String, Optional ByVal bolSaveAsInspection As Boolean = False)
            Dim IDs As New Collection
            Dim delIDs As New Collection
            'Dim tnkID As Integer
            Dim index As Integer
            Dim oTempInfo As MUSTER.Info.CompartmentInfo
            Try
                'Dim colCompartmentContained As MUSTER.Info.CompartmentCollection
                'RaiseEvent evtTankInfoCompCol(colCompartmentContained)
                For Each oTempInfo In oTankInfo.CompartmentCollection.Values
                    If oTempInfo.IsDirty Then
                        oCompartmentInfo = oTempInfo
                        If oCompartmentInfo.COMPARTMENTNumber < 0 Then
                            If oCompartmentInfo.Deleted Then
                                delIDs.Add(oCompartmentInfo.ID)
                            Else
                                IDs.Add(oCompartmentInfo.ID)
                                ' sorty array in ascending order to save the order in which compartments were entered
                            End If
                        Else
                            If oCompartmentInfo.Deleted Then
                                delIDs.Add(oCompartmentInfo.ID)
                            End If
                            If oCompartmentInfo.COMPARTMENTNumber <= 0 Then
                                oCompartmentInfo.CreatedBy = UserID
                            Else
                                oCompartmentInfo.ModifiedBy = UserID
                            End If
                            Me.Save(moduleID, staffID, returnVal, UserID, True, , bolSaveAsInspection)
                        End If
                    ElseIf oTankInfo.PipesDirty Then
                        oCompartmentInfo = oTempInfo
                        oCompartmentPipe.Flush(moduleID, staffID, returnVal, UserID, bolSaveAsInspection)
                        If Not returnVal = String.Empty Then
                            Exit Sub
                        End If
                    End If
                Next
                If Not (delIDs Is Nothing) Then
                    For index = 1 To delIDs.Count
                        oTempInfo = oTankInfo.CompartmentCollection.Item(CType(delIDs.Item(index), String))
                        oTankInfo.CompartmentCollection.Remove(oTempInfo)
                        'RaiseEvent evtCompInfoTank(oTempInfo, "REMOVE")
                    Next
                End If
                If Not (IDs Is Nothing) Then
                    Dim sortedList As New sortedList
                    For index = 0 To IDs.Count - 1
                        Dim strArr() As String = CType(IDs(index + 1), String).Split("|")
                        sortedList.Add(CType(strArr(1), Integer), CType(strArr(0), Integer))
                    Next
                    For index = sortedList.Count - 1 To 0 Step -1
                        oCompartmentInfo = oTankInfo.CompartmentCollection.Item(CType(sortedList.GetByIndex(index), String) + "|" + CType(sortedList.GetKey(index), String))
                        If oTankInfo.CompartmentCollection.Count > 1 And _
                            oCompartmentInfo.COMPARTMENTNumber < 0 And _
                            oCompartmentInfo.Capacity = 0 And _
                            oCompartmentInfo.Substance = 0 And _
                            oCompartmentInfo.CCERCLA = 0 And _
                            oCompartmentInfo.FuelTypeId = 0 Then
                            oTankInfo.CompartmentCollection.Remove(oCompartmentInfo)
                            IDs.Remove(index + 1)
                        Else
                            If oCompartmentInfo.COMPARTMENTNumber <= 0 Then
                                oCompartmentInfo.CreatedBy = UserID
                            Else
                                oCompartmentInfo.ModifiedBy = UserID
                            End If
                            Me.Save(moduleID, staffID, returnVal, UserID, True)
                        End If
                    Next
                    For index = 1 To IDs.Count
                        Dim colKey As String = CType(IDs.Item(index), String)
                        oTempInfo = oTankInfo.CompartmentCollection.Item(colKey)
                        oTankInfo.CompartmentCollection.ChangeKey(colKey, oTempInfo.ID)
                        'RaiseEvent evtCompartmentChangeKey(colKey, oTempInfo.ID)
                    Next
                End If
                'oCompartmentPipe.Flush()
                RaiseEvent evtCompartmentsChanged(oCompartmentInfo.IsDirty)
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
            'Dim colCompartmentContained As MUSTER.Info.CompartmentCollection
            'RaiseEvent evtTankInfoCompCol(colCompartmentContained)
            Dim strArr() As String = oTankInfo.CompartmentCollection.GetKeys()
            colIndex = Array.BinarySearch(strArr, Me.ID.ToString)
            If colIndex + direction > -1 And _
                colIndex + direction <= strArr.GetUpperBound(0) Then
                Return oTankInfo.CompartmentCollection.Item(strArr.GetValue(colIndex + direction)).ID.ToString
            Else
                Return oTankInfo.CompartmentCollection.Item(strArr.GetValue(colIndex)).ID.ToString
            End If
            If colIndex + direction > -1 Then
                If colIndex + direction <= strArr.GetUpperBound(0) Then 'colIndex + direction <= nArr.GetUpperBound(0) Then
                    Return oTankInfo.CompartmentCollection.Item(strArr.GetValue(colIndex + direction)).ID.ToString
                Else
                    Return oTankInfo.CompartmentCollection.Item(strArr.GetValue(0)).ID.ToString
                End If
            Else
                Return oTankInfo.CompartmentCollection.Item(strArr.GetValue(strArr.GetUpperBound(0))).ID.ToString
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oCompartmentInfo = New MUSTER.Info.CompartmentInfo
        End Sub
        Public Sub Reset()
            Dim IDs As New Collection
            'Dim colCompartmentContained As MUSTER.Info.CompartmentCollection
            'RaiseEvent evtTankInfoCompCol(colCompartmentContained)
            Dim xCompartmentInfo As MUSTER.Info.CompartmentInfo
            If oTankInfo.CompartmentCollection.Count > 0 Then
                For Each xCompartmentInfo In oTankInfo.CompartmentCollection.Values
                    If xCompartmentInfo.COMPARTMENTNumber < 0 Then
                        IDs.Add(xCompartmentInfo.ID)
                    ElseIf xCompartmentInfo.IsDirty Then
                        xCompartmentInfo.Reset()
                    End If
                Next
            Else
                xCompartmentInfo.Reset()
            End If
            If Not (IDs Is Nothing) Then
                For index As Integer = 1 To IDs.Count
                    xCompartmentInfo = oTankInfo.CompartmentCollection.Item(CType(IDs.Item(index), String))
                    oTankInfo.CompartmentCollection.Remove(xCompartmentInfo)
                Next
            End If
        End Sub
#End Region
#Region "Miscellaneous Operations"
        'Returns a datatable of the Compartments in the collection
        Public Function EntityTable() As DataTable

            Dim oCompartmentInfoLocal As MUSTER.Info.CompartmentInfo
            Dim dr As DataRow
            Dim tbCompartmentTable As New DataTable

            Try

                tbCompartmentTable.Columns.Add("ID", GetType(String))
                tbCompartmentTable.Columns.Add("TANK_ID", GetType(Integer))
                tbCompartmentTable.Columns.Add("COMPARTMENT NUMBER", GetType(Integer))
                tbCompartmentTable.Columns.Add("COMPARTMENT #", GetType(Integer))
                tbCompartmentTable.Columns.Add("CAPACITY", GetType(Integer))
                tbCompartmentTable.Columns.Add("SUBSTANCE", GetType(Integer))
                tbCompartmentTable.Columns.Add("CERCLA#", GetType(Integer))
                tbCompartmentTable.Columns.Add("FUEL TYPE ID", GetType(Integer))
                tbCompartmentTable.Columns.Add("MANIFOLD INFO", GetType(String))
                tbCompartmentTable.Columns.Add("Deleted", GetType(Boolean))
                tbCompartmentTable.Columns.Add("Created By", GetType(String))
                tbCompartmentTable.Columns.Add("Created On", GetType(DateTime))
                tbCompartmentTable.Columns.Add("Modified By", GetType(String))
                tbCompartmentTable.Columns.Add("Modified On", GetType(DateTime))

                'Dim colCompartmentContained As MUSTER.Info.CompartmentCollection
                'RaiseEvent evtTankInfoCompCol(colCompartmentContained)
                For Each oCompartmentInfoLocal In oTankInfo.CompartmentCollection.Values
                    dr = tbCompartmentTable.NewRow()
                    dr("ID") = oCompartmentInfoLocal.ID
                    dr("TANK_ID") = oCompartmentInfoLocal.TankId
                    dr("COMPARTMENT NUMBER") = oCompartmentInfoLocal.COMPARTMENTNumber
                    dr("COMPARTMENT #") = oCompartmentInfoLocal.COMPARTMENTNumber
                    dr("CAPACITY") = oCompartmentInfoLocal.Capacity
                    dr("CERCLA#") = oCompartmentInfoLocal.CCERCLA
                    dr("SUBSTANCE") = oCompartmentInfoLocal.Substance
                    dr("FUEL TYPE ID") = oCompartmentInfoLocal.FuelTypeId
                    dr("MANIFOLD INFO") = "-NA-"
                    dr("Deleted") = oCompartmentInfoLocal.Deleted
                    dr("Created By") = oCompartmentInfoLocal.CreatedBy
                    dr("Created On") = oCompartmentInfoLocal.CreatedOn
                    dr("Modified By") = oCompartmentInfoLocal.ModifiedBy
                    dr("Modified On") = oCompartmentInfoLocal.ModifiedOn
                    tbCompartmentTable.Rows.Add(dr)
                Next

                Return tbCompartmentTable
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Function GetDataSet(ByVal strSQL As String) As DataSet
            Try
                Dim ds As DataSet
                ds = oCompartmentDB.DBGetDS(strSQL)
                Return ds
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Sub ChangeCompartmentNumberKey(Optional ByVal newCompNum As Integer = 0, Optional ByVal newTnkID As Integer = 0, Optional ByRef cmpInfo As MUSTER.Info.CompartmentInfo = Nothing)
            Dim oCompartmentInfoLocal As MUSTER.Info.CompartmentInfo
            If cmpInfo Is Nothing Then
                cmpInfo = oCompartmentInfo
            End If
            Dim OldKey() As String = cmpInfo.ID.Split("|")
            Try
                If newCompNum = 0 Then
                    cmpInfo.ID = newTnkID.ToString + "|" + OldKey(1)
                    cmpInfo.TankId = newTnkID
                    Exit Try
                End If
                If newTnkID = 0 Then
                    cmpInfo.ID = OldKey(0) + "|" + newCompNum.ToString
                    cmpInfo.COMPARTMENTNumber = newCompNum
                    Exit Try
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
            oTankInfo.CompartmentCollection.ChangeKey(OldKey(0) + "|" + OldKey(1), cmpInfo.ID)
            'RaiseEvent evtCompartmentChangeKey(OldKey(0) + "|" + OldKey(1), cmpInfo.ID)
        End Sub
        Public Function getManifold(ByVal nTank_ID As Integer) As DataTable
            Dim ds As DataSet
            Try
                getManifold = Nothing
                ds = oCompartmentDB.getManifold(nTank_ID)
                Return ds.Tables(0)
            Catch ex As Exception
                Throw ex
            End Try
        End Function
        Public Function CompartmentsTable(ByVal tnkID As Integer) As DataTable
            Dim oCompInfoLocal As MUSTER.Info.CompartmentInfo
            Dim dr As DataRow
            Dim dtCompTable As New DataTable
            Try
                dtCompTable.Columns.Add("Tank_ID", Type.GetType("System.Int64"))
                dtCompTable.Columns.Add("Substance", Type.GetType("System.String"))
                dtCompTable.Columns.Add("Capacity", Type.GetType("System.Int64"))

                'Dim colCompartmentContained As MUSTER.Info.CompartmentCollection
                'RaiseEvent evtTankInfoCompCol(colCompartmentContained)
                For Each oCompInfoLocal In oTankInfo.CompartmentCollection.Values
                    If oCompInfoLocal.TankId = tnkID And Not (oCompInfoLocal.Deleted) Then
                        dr = dtCompTable.NewRow()
                        dr("Tank_ID") = oCompInfoLocal.TankId
                        dr("Substance") = IIf(oProperty.Retrieve(oCompInfoLocal.Substance).Name Is Nothing, String.Empty, oProperty.Retrieve(oCompInfoLocal.Substance).Name) ' return string desc
                        dr("Capacity") = oCompInfoLocal.Capacity
                        dtCompTable.Rows.Add(dr)
                    End If
                Next
                Return dtCompTable
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
#End Region
#Region "External Event Handlers"

        'Private Sub PipeCommentsChanged(ByVal bolValue As Boolean) Handles oCompartmentPipe.evtPipeCommentsChanged
        '    RaiseEvent evtPipeCommentsChanged(bolValue)
        'End Sub
        'Private Sub checkCAPStatusPipe(ByVal nVal As Integer) Handles oCompartmentPipe.evtPipeCAPChanged
        '    RaiseEvent evtCAPStatusPipe(nVal)
        'End Sub
#End Region
#End Region
#Region "Event Handlers"
        Private Sub CompartmentChanged(ByVal bolValue As Boolean) Handles oCompartmentInfo.CompartmentInfoChanged
            RaiseEvent evtCompartmentChanged(bolValue)
        End Sub
        'added by kiran
        'Private Sub CompartmentPipeCol(ByVal TnkID As String, ByVal pipeCol As MUSTER.Info.PipesCollection) Handles oCompartmentPipe.evtPipeColCompartment
        '    'Dim oCompartmentInfoLocal As MUSTER.Info.CompartmentInfo
        '    'Try

        '    '    'oCompartmentInfoLocal = colCompartment.Item(compID)
        '    '    'If Not (oCompartmentInfoLocal Is Nothing) Then
        '    '    '    oCompartmentInfoLocal.pipesCollection = pipeCol
        '    '    'End If
        '    'Catch ex As Exception
        '    '    If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '    '    Throw ex
        '    'End Try
        '    RaiseEvent evtPipeColTank(TnkID, pipeCol)
        'End Sub
        'Private Sub pipecommentsCol(ByVal pipeID As Integer, ByVal compID As Integer, ByVal commentsCol As MUSTER.Info.CommentsCollection) Handles oCompartmentPipe.evtPipesCommentsCol
        '    RaiseEvent evtPipeCommentsCol(pipeID, compID, oCompartmentInfo.TankId, commentsCol)
        'End Sub
        ''end changes
        'Private Sub PipeInfoCompartment(ByVal pipeInfo As MUSTER.Info.PipeInfo, ByVal strDesc As String) Handles oCompartmentPipe.evtPipeInfoCompartment
        '    RaiseEvent evtPipeInfoCompartment(pipeInfo, strDesc)
        'End Sub
        'Private Sub CompInfoPipeCol(ByRef colPipe As MUSTER.Info.PipesCollection) Handles oCompartmentPipe.evtCompInfoPipeCol
        '    colPipe = oCompartmentInfo.pipesCollection
        'End Sub
        'Private Sub PipeChangeKey(ByVal oldID As String, ByVal newID As String) Handles oCompartmentPipe.evtPipeChangeKey
        '    Try
        '        If oCompartmentInfo.pipesCollection.Contains(oldID) Then
        '            oCompartmentInfo.pipesCollection.ChangeKey(oldID, newID)
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'Private Sub SyncPipeInCol(ByVal pipeInfo As MUSTER.Info.PipeInfo) Handles oCompartmentPipe.evtSyncPipeInCol
        '    RaiseEvent evtSyncPipeInCol(pipeInfo)
        'End Sub
        Private Sub PipeFacCapStatus(ByVal facID As Integer) Handles oCompartmentPipe.evtFacCapStatus
            RaiseEvent evtFacCapStatus(facID)
        End Sub
#End Region
    End Class
End Namespace
