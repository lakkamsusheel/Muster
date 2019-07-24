'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.pReportParams
'   Provides the collection and info object to the client for manipulating
'     Report Parameters.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'   1.0         AN      12/02/04    Added documentation.
'   1.0         AN      12/01/04    Original class definition.
'   1.1         AN      01/03/05    Added Try catch and Exception Handling/Logging
'
' Operations
' Function                      Description
' New()                         Initializes the ProfileCollection and ProfileInfo objects.
' Retrieve(Key, [ShowDeleted])  Sets the internal ProfileInfo to the ProfileInfo matching the 
'                                supplied key.  In the event that a partial key is supplied,
'                                all matching ProfileInfos are populated to the internal 
'                                ProfileCollection and the internal ProfileInfo object is
'                                set to the first member of the collection.  In either case,
'                                the internal ProfileInfo object is returned to the client.
' Add(ProfileInfo)              Adds the supplied ProfileInfo object to the internal ProfileCollection
'                                and sets the internal ProfileInfo object to same.
' Remove(ProfileInfo)           Removes the supplied ProfileInfo object from the internal ProfileCollection
'                                if it is contained by the collection.
' Items()                       Returns the internal ProfileCollection (a Dictionary object) to the client.
' Values()                      Returns the collection of ProfileInfo objects to the client (used in for..next).
' Clear()                       Sets the internal ProfileInfo object to an empty ProfileInfo and clears
'                                the internal ProfileCollection.
' Reset()                       Reverts the internal ProfileInfo object to it's state when last retrieved
'                                from or marshalled to the repository.
' Save()                        Marshalls the internal ProfileInfo to the repository.
' Flush()                       Marshalls all modified/new ProfileInfo objects in the ProfileCollection
'                                to the repository.
'
' Properties
' Name                      Description
'  ReportName                   Gets or Sets the FileName of the Report
'  Param                        Gets or Sets the name of the Parameter of the current item
'  ParamDescription             Gets or Sets the text of the parameter Description
'  ColIsDirty                   Returns a BOOLEAN if the Collections of Params is Dirty
'  CreatedBy                    ReadOnly returns name of user that created the item
'  CreatedOn                    ReadOnly returns the date the item was created
'  ModifiedBy                   ReadOnly returns the name of the last person to modify the item
'  ModifiedOn                   ReadOnly returns the date the item was last modified
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pReportParams
#Region "Private Member Variables"
        Private colParams As Muster.Info.ProfileCollection
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private WithEvents oParamInfo As Muster.Info.ProfileInfo
        Private oProfileDB As New Muster.DataAccess.ProfileDB
        Private strParent As String
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Public Events"
        Public Event ReportParamChanged(ByVal bolValue As Boolean)
#End Region
#Region "Constructors"
        Public Sub New()
            oParamInfo = New Muster.Info.ProfileInfo
            colParams = New Muster.Info.ProfileCollection
        End Sub
        Public Sub New(ByVal ReportName As String)
            oParamInfo = New Muster.Info.ProfileInfo
            colParams = New Muster.Info.ProfileCollection
            Me.ReportID = ReportName
        End Sub
#End Region
#Region "Exposed Attributes"
        Public ReadOnly Property ProfileID() As String
            Get
                Return oParamInfo.ID
            End Get
        End Property
        Public Property ReportID() As String
            Get
                Return strParent
            End Get
            Set(ByVal Value As String)
                strParent = Value
                Me.colParams.Clear()
                Me.Retrieve("SYSTEM|REPORTPARAMS|" & Value, False)
            End Set
        End Property

        Public Property Param() As String
            Get
                Return oParamInfo.ProfileMod2
            End Get
            Set(ByVal Value As String)
                oParamInfo.ProfileMod2 = Value
                colParams(oParamInfo.ID) = oParamInfo
            End Set
        End Property

        Public Property ParamDescription() As String
            Get
                If Not oParamInfo Is Nothing Then
                    Return oParamInfo.ProfileValue
                Else
                    Return String.Empty
                End If
            End Get

            Set(ByVal value As String)
                oParamInfo.ProfileValue = value
                colParams(oParamInfo.ID) = oParamInfo
            End Set
        End Property

        Public Property colIsDirty() As Boolean
            Get
                Dim xProfInf As Muster.Info.ProfileInfo
                For Each xProfInf In colParams.Values
                    If xProfInf.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
            Set(ByVal Value As Boolean)

            End Set
        End Property

        Public ReadOnly Property CreatedBy() As String
            Get
                Return oParamInfo.CreatedBy
            End Get
        End Property

        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oParamInfo.CreatedOn
            End Get
        End Property

        Public ReadOnly Property ModifiedBy() As String
            Get
                Return oParamInfo.ModifiedBy
            End Get
        End Property

        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oParamInfo.ModifiedOn
            End Get
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        ' Obtains and returns an entity as called for by name
        Public Function Retrieve(ByVal FullKey As String, Optional ByVal ShowDeleted As Boolean = False) As Muster.Info.ProfileInfo
            oParamInfo = Nothing
            Try
                If colParams.Contains(FullKey) Then
                    oParamInfo = colParams.Item(FullKey)
                    Return oParamInfo
                Else
                    Dim strArray() As String
                    Dim colTemp As Muster.Info.ProfileCollection
                    Dim oInfTemp As Muster.Info.ProfileInfo
                    strArray = FullKey.Split("|")
                    colTemp = oProfileDB.DBGetByKey(strArray, ShowDeleted)
                    For Each oInfTemp In colTemp.Values
                        colParams.Add(oInfTemp)
                        oParamInfo = oInfTemp
                    Next
                    'If colParams.Count > 0 Then
                    '    strArray = colParams.GetKeys()
                    '    Return Retrieve(strArray(0))
                    'End If
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function

        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try
                oProfileDB.Put(oParamInfo, moduleID, staffID, returnVal)
                If Not returnVal = String.Empty Then
                    Exit Sub
                End If

                oParamInfo.Archive()
                oParamInfo.IsDirty = False
                RaiseEvent ReportParamChanged(oParamInfo.IsDirty)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
#End Region
#Region "Collection Operations"
        Function GetAll() As Muster.Info.ProfileCollection
            Try
                colParams.Clear()
                colParams = oProfileDB.GetAllInfo
                Return colParams
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oProfileInf As Muster.Info.ProfileInfo)

            Try
                oParamInfo = oProfileInf
                colParams.Add(oParamInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oProfileInf As Muster.Info.ProfileInfo)

            Try
                colParams.Remove(oProfileInf)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

            Throw New Exception("Profile Info " & oProfileInf.ID & " is not in the collection of profile data.")

        End Sub
        Public Function Items() As Muster.Info.ProfileCollection
            Return colParams
        End Function
        Public Function Values() As ICollection
            Return colParams.Values
        End Function
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim xProfInf As MUSTER.Info.ProfileInfo
            For Each xProfInf In colParams.Values
                If xProfInf.IsDirty Then
                    oParamInfo = xProfInf
                    Me.Save(moduleID, staffID, returnVal)
                End If
            Next
            'Adam Nall - Added to change the pProfile isDirty. It changes this on the save 
            '            of the single profileInfo but not the parent class isDirty
            Me.colIsDirty = False
        End Sub
#End Region
#Region "General Operations"
        Public Sub Clear()
            oParamInfo = New MUSTER.Info.ProfileInfo
            colParams.Clear()
        End Sub
        Public Sub Reset()
            If Not oParamInfo Is Nothing Then
                oParamInfo.Reset()
            End If
            RaiseEvent ReportParamChanged(oParamInfo.IsDirty)
        End Sub
#End Region
#Region "Miscellaneous Operations"

#End Region
#End Region
#Region "External Event Handlers"
        Private Sub ParamChanged(ByVal bolValue As Boolean) Handles oParamInfo.InfoBecameDirty
            RaiseEvent ReportParamChanged(bolValue)
        End Sub
#End Region
    End Class
End Namespace
