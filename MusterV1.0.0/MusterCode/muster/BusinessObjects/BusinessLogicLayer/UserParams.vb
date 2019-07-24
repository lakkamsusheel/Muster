''-------------------------------------------------------------------------------
'' MUSTER.BusinessLogic.pUserParams
''   Provides the collection and info object to the client for manipulating
''     User Parameters.
''
'' Copyright (C) 2004 CIBER, Inc.
'' All rights reserved.
''
'' Release   Initials    Date        Description
''   1.0         PN      12/05/04    Original class definition.
''
'' Operations
'' Function                      Description
'' New()                         Initializes the ProfileCollection and ProfileInfo objects.
'' Retrieve(Key, [ShowDeleted])  Sets the internal ProfileInfo to the ProfileInfo matching the 
''                                supplied key.  In the event that a partial key is supplied,
''                                all matching ProfileInfos are populated to the internal 
''                                ProfileCollection and the internal ProfileInfo object is
''                                set to the first member of the collection.  In either case,
''                                the internal ProfileInfo object is returned to the client.
'' Add(ProfileInfo)              Adds the supplied ProfileInfo object to the internal ProfileCollection
''                                and sets the internal ProfileInfo object to same.
'' Remove(ProfileInfo)           Removes the supplied ProfileInfo object from the internal ProfileCollection
''                                if it is contained by the collection.
'' Items()                       Returns the internal ProfileCollection (a Dictionary object) to the client.
'' Values()                      Returns the collection of ProfileInfo objects to the client (used in for..next).
'' Clear()                       Sets the internal ProfileInfo object to an empty ProfileInfo and clears
''                                the internal ProfileCollection.
'' Reset()                       Reverts the internal ProfileInfo object to it's state when last retrieved
''                                from or marshalled to the repository.
'' Save()                        Marshalls the internal ProfileInfo to the repository.
'' Flush()                       Marshalls all modified/new ProfileInfo objects in the ProfileCollection
''                                to the repository.
''
'' Properties
'' Name                      Description
''  ReportName                   Gets or Sets the FileName of the Report
''  Param                        Gets or Sets the name of the Parameter of the current item
''  ParamDescription             Gets or Sets the text of the parameter Description
''  ColIsDirty                   Returns a BOOLEAN if the Collections of Params is Dirty
''  CreatedBy                    ReadOnly returns name of user that created the item
''  CreatedOn                    ReadOnly returns the date the item was created
''  ModifiedBy                   ReadOnly returns the name of the last person to modify the item
''  ModifiedOn                   ReadOnly returns the date the item was last modified
''-------------------------------------------------------------------------------



'Namespace MUSTER.BusinessLogic
'    <Serializable()> _
'    Public Class pUserParams
'#Region "Private Member Variables"
'        Private colUsers As Muster.Info.UserCollection
'        Private strGroupOwnerName As String
'        Private colIndex As Int64 = 0
'        Private colKey As String = String.Empty
'        Private oUserInfo As Muster.Info.UserInfo
'        Private intManagerID As Integer
'        Private oUserDB As New Muster.DataAccess.UserDB

'#End Region
'#Region "Constructors"
'        Public Sub New()
'            oUserInfo = New Muster.Info.UserInfo
'            colUsers = New Muster.Info.UserCollection
'        End Sub
'        Public Sub New(ByVal strUserID As String)
'            oUserInfo = New Muster.Info.UserInfo
'            colUsers = New Muster.Info.UserCollection
'            Me.ManagerID = strUserID
'        End Sub
'#End Region
'#Region "Exposed Attributes"
'        Public Property ManagerID() As String
'            Get
'                Return intManagerID
'            End Get
'            Set(ByVal Value As String)
'                intManagerID = Value
'                Me.colUsers.Clear()
'            End Set
'        End Property
'        Public Property UserName() As String
'            Get
'                Return strParent
'            End Get
'            Set(ByVal Value As String)
'                strParent = Value
'                Me.colParams.Clear()
'                Me.Retrieve("SYSTEM|REPORTPARAMS|" & Value, False)
'            End Set
'        End Property

'        Public Property Param() As String
'            Get
'                Return oParamInfo.ProfileMod2
'            End Get
'            Set(ByVal Value As String)
'                oParamInfo.ProfileMod2 = Value
'                colParams(oParamInfo.ID) = oParamInfo
'            End Set
'        End Property

'        Public Property ParamDescription() As String
'            Get
'                If Not oParamInfo Is Nothing Then
'                    Return oParamInfo.ProfileValue
'                Else
'                    Return String.Empty
'                End If
'            End Get

'            Set(ByVal value As String)
'                oParamInfo.ProfileValue = value
'                colParams(oParamInfo.ID) = oParamInfo
'            End Set
'        End Property

'        Public Property colIsDirty() As Boolean
'            Get
'                Dim xProfInf As Muster.Info.ProfileInfo
'                For Each xProfInf In colParams.Values
'                    If xProfInf.IsDirty Then
'                        Return True
'                        Exit Property
'                    End If
'                Next
'                Return False
'            End Get
'            Set(ByVal Value As Boolean)

'            End Set
'        End Property

'        Public ReadOnly Property CreatedBy() As String
'            Get
'                Return oParamInfo.CreatedBy
'            End Get
'        End Property

'        Public ReadOnly Property CreatedOn() As Date
'            Get
'                Return oParamInfo.CreatedOn
'            End Get
'        End Property

'        Public ReadOnly Property ModifiedBy() As String
'            Get
'                Return oParamInfo.ModifiedBy
'            End Get
'        End Property

'        Public ReadOnly Property ModifiedOn() As Date
'            Get
'                Return oParamInfo.ModifiedOn
'            End Get
'        End Property
'#End Region
'#Region "Exposed Operations"
'#Region "Info Operations"
'        ' Obtains and returns an entity as called for by name
'        Public Function Retrieve(ByVal Key As Integer) As Muster.Info.UserInfo
'            oUserInfo = Nothing
'            Try
'                If colUsers.Contains(Key) Then
'                    oUserInfo = colUsers.Item(Key)
'                    Return oUserInfo
'                Else

'                    Dim colTemp As Muster.Info.UserCollection
'                    Dim oInfTemp As Muster.Info.UserInfo

'                    colTemp = oProfileDB.DBGetByKey(strArray, ShowDeleted)
'                    For Each oInfTemp In colTemp.Values
'                        colParams.Add(oInfTemp)
'                    Next
'                    'If colParams.Count > 0 Then
'                    '    strArray = colParams.GetKeys()
'                    '    Return Retrieve(strArray(0))
'                    'End If
'                End If
'            Catch ex As Exception
'                Throw ex
'            End Try

'        End Function

'        Public Sub Save()
'            oProfileDB.Put(oParamInfo)
'            oParamInfo.Archive()
'            oParamInfo.IsDirty = False
'        End Sub
'#End Region
'#Region "Collection Operations"
'        Function GetAll() As Muster.Info.ProfileCollection
'            colParams.Clear()
'            colParams = oProfileDB.GetAllInfo
'            Return colParams
'        End Function
'        'Adds an entity to the collection as supplied by the caller
'        Public Sub Add(ByRef oProfileInf As Muster.Info.ProfileInfo)

'            Try
'                oParamInfo = oProfileInf
'                colParams.Add(oParamInfo)
'            Catch ex As Exception
'                Throw ex
'            End Try

'        End Sub
'        'Removes the entity supplied from the collection
'        Public Sub Remove(ByVal oProfileInf As Muster.Info.ProfileInfo)

'            Try
'                colParams.Remove(oProfileInf)
'            Catch ex As Exception
'                Throw ex
'            End Try

'            Throw New Exception("Profile Info " & oProfileInf.ID & " is not in the collection of profile data.")

'        End Sub
'        Public Function Items() As Muster.Info.ProfileCollection
'            Return colParams
'        End Function

'        Public Function Values() As ICollection
'            Return colParams.Values
'        End Function

'        Public Sub Flush()
'            Dim xProfInf As Muster.Info.ProfileInfo
'            For Each xProfInf In colParams.Values
'                If xProfInf.IsDirty Then
'                    oParamInfo = xProfInf
'                    Me.Save()
'                End If
'            Next
'            'Adam Nall - Added to change the pProfile isDirty. It changes this on the save 
'            '            of the single profileInfo but not the parent class isDirty
'            Me.colIsDirty = False
'        End Sub
'#End Region
'#Region "General Operations"
'        Public Function Clear()
'            oParamInfo = New Muster.Info.ProfileInfo
'            colParams.Clear()
'        End Function
'        Public Function Reset()
'            oParamInfo.Reset()
'        End Function
'#End Region
'#Region "Miscellaneous Operations"

'#End Region
'#End Region
'    End Class
'End Namespace
