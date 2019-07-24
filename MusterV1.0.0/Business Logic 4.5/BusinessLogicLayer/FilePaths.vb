'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.pFilePath
'   Provides the collection and info object to the client for manipulating
'     profileinfo data.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'   1.0         PN      12/11/04    Original class definition.
'               PN      12/22/04    Added function FilepathTable  
'   1.1         JC      12/28/04    Added event for data update notifications.
'                                   Added event firing to SAVE()
'                                   Added event firing to FLUSH()
'   1.2         AN      01/03/05    Added Try catch and Exception Handling/Logging
'   1.3         MNR     01/25/05    Added GetNext(), GetPrevious() and GetNextPrev(..) functions
'
' Operations
' Function                      Description
' New()                         Initializes the ProfileCollection and profileinfo objects.
' Retrieve(Key, [ShowDeleted])  Sets the internal profileinfo to the profileinfo matching the 
'                                supplied key.  In the event that a partial key is supplied,
'                                all matching profileinfos are populated to the internal 
'                                ProfileCollection and the internal profileinfo object is
'                                set to the first member of the collection.  In either case,
'                                the internal profileinfo object is returned to the client.
' Save()                        Marshalls the internal profileinfo to the repository.
' Add(profileinfo)              Adds the supplied profileinfo object to the internal ProfileCollection
'                                and sets the internal profileinfo object to same.
' Remove(profileinfo)           Removes the supplied profileinfo object from the internal ProfileCollection
'                                if it is contained by the collection.
' Items()                       Returns the internal ProfileCollection (a Dictionary object) to the client.
' Values()                      Returns the collection of profileinfo objects to the client (used in for..next).
' Flush()                       Marshalls all modified/new profileinfo objects in the ProfileCollection
'                                to the repository.
' Clear()                       Sets the internal profileinfo object to an empty profileinfo and clears
'                                the internal ProfileCollection.
' Reset()                       Reverts all the internal profileinfo object to it's state when last retrieved
'                                from or marshalled to the repository.
' FilePathTable()               Returns FileName and FilePath of all the files in the collection
' 
'-------------------------------------------------------------------------------
'
' TODO - 12/29 - Check Operations list and add Attributes list to header.
'

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pFilePaths
#Region "Private Member Variables"
        Private colFilePaths As Muster.Info.ProfileCollection
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private WithEvents oFilePathInfo As Muster.Info.ProfileInfo
        Private oFilePathDB As New Muster.DataAccess.ProfileDB
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Public Events"
        Public Event FilePathIsDirty(ByVal IsDirty As Boolean)
#End Region
#Region "Constructors"
        Public Sub New()
            oFilePathInfo = New Muster.Info.ProfileInfo
            colFilePaths = New Muster.Info.ProfileCollection
        End Sub
#End Region
#Region "Exposed Attributes"

        Public Property FilePaths() As MUSTER.Info.ProfileCollection
            Get

                If colFilePaths Is Nothing Then
                    colFilePaths = New MUSTER.Info.ProfileCollection
                End If
                Return colFilePaths
            End Get
            Set(ByVal Value As MUSTER.Info.ProfileCollection)
                colFilePaths = Value
            End Set
        End Property

        Public Property FilePathInfo() As MUSTER.Info.ProfileInfo
            Get

                If oFilePathInfo Is Nothing Then
                    oFilePathInfo = New MUSTER.Info.ProfileInfo
                End If
                Return oFilePathInfo
            End Get
            Set(ByVal Value As MUSTER.Info.ProfileInfo)
                oFilePathInfo = Value
            End Set
        End Property

        Public Property ID() As String
            Get
                Return FilePathInfo.ID
            End Get

            Set(ByVal value As String)
                FilePathInfo.ID = value
            End Set
        End Property

        Public Property FileKey() As String
            Get
                Return FilePathInfo.ProfileKey
            End Get
            Set(ByVal Value As String)
                FilePathInfo.ProfileKey = Value
                FilePaths(oFilePathInfo.ID) = oFilePathInfo
            End Set
        End Property

        Public Property Name() As String
            Get
                Return FilePathInfo.ProfileMod1
            End Get
            Set(ByVal Value As String)
                FilePathInfo.ProfileMod1 = Value
                FilePaths(oFilePathInfo.ID) = oFilePathInfo
            End Set
        End Property

        Public Property Name2() As String
            Get
                Return FilePathInfo.ProfileMod2
            End Get
            Set(ByVal Value As String)

                FilePathInfo.ProfileMod2 = Value
                FilePaths(oFilePathInfo.ID) = oFilePathInfo
            End Set
        End Property

        Public Property FilePath() As String
            Get
                Return FilePathInfo.ProfileValue
            End Get

            Set(ByVal value As String)
                FilePathInfo.ProfileValue = value
                FilePaths.Add(oFilePathInfo)
            End Set
        End Property

        Public Property Deleted() As Boolean
            Get
                Return FilePathInfo.Deleted
            End Get

            Set(ByVal value As Boolean)
                FilePathInfo.Deleted = value
                FilePaths(oFilePathInfo.ID) = oFilePathInfo
            End Set
        End Property

        Public Property colIsDirty() As Boolean
            Get
                Dim xFilePathInf As MUSTER.Info.ProfileInfo
                For Each xFilePathInf In FilePaths.Values
                    If xFilePathInf.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
            Set(ByVal Value As Boolean)

            End Set
        End Property

        Public Property IsDirty() As Boolean
            Get
                Return FilePathInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                FilePathInfo.IsDirty = value
                FilePaths.Add(oFilePathInfo)
            End Set
        End Property

        Public ReadOnly Property CreatedBy() As String
            Get
                Return FilePathInfo.CreatedBy
            End Get
        End Property

        Public ReadOnly Property CreatedOn() As Date
            Get
                Return FilePathInfo.CreatedOn
            End Get
        End Property

        Public ReadOnly Property ModifiedBy() As String
            Get
                Return FilePathInfo.ModifiedBy
            End Get
        End Property

        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return FilePathInfo.ModifiedOn
            End Get
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        ' Obtains and returns an entity as called for by name
        Public Function Retrieve(ByVal FullKey As String, Optional ByVal ShowDeleted As Boolean = False) As MUSTER.Info.ProfileInfo
            Try
                If FilePaths.Contains(FullKey) Then
                    FilePathInfo = FilePaths.Item(FullKey)
                    Return FilePathInfo
                Else
                    Dim strArray() As String
                    Dim colTemp As MUSTER.Info.ProfileCollection
                    Dim oInfTemp As MUSTER.Info.ProfileInfo
                    strArray = FullKey.Split("|")
                    colTemp = oFilePathDB.DBGetByKey(strArray, ShowDeleted)
                    For Each oInfTemp In colTemp.Values
                        colFilePaths.Add(oInfTemp)
                    Next

                    If colTemp.Values.Count > 0 Then
                        FilePathInfo = FilePaths.Item(FullKey)
                        Return oFilePathInfo
                    End If




                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        'Saves the data in the current Info object
        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try
                oFilePathDB.Put(oFilePathInfo, moduleID, staffID, returnVal)
                If Not returnVal = String.Empty Then
                    Exit Sub
                End If

                FilePathInfo.Archive()
                FilePathInfo.IsDirty = False
                RaiseEvent FilePathIsDirty(Me.colIsDirty)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
#End Region
#Region "Collection Operations"
        Function GetAll() As MUSTER.Info.ProfileCollection
            Try
                FilePaths.Clear()
                FilePaths = oFilePathDB.GetAllInfo
                Return FilePaths
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oProfileInf As MUSTER.Info.ProfileInfo)

            Try
                FilePathInfo = oProfileInf
                FilePaths.Add(oFilePathInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        Public Sub Remove(ByVal oProfileInf As MUSTER.Info.ProfileInfo)

            Try
                FilePaths.Remove(oProfileInf)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

            Throw New Exception("Profile Info " & oProfileInf.ID & " is not in the collection of profile data.")

        End Sub
        Public Function Items() As MUSTER.Info.ProfileCollection
            Return FilePaths
        End Function
        Public Function Values() As ICollection
            Return FilePaths.Values
        End Function
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String)
            Dim xFilePathInf As MUSTER.Info.ProfileInfo
            For Each xFilePathInf In FilePaths.Values
                If xFilePathInf.IsDirty Then
                    FilePathInfo = xFilePathInf
                    If FilePathInfo.User = "SYSTEM" Then
                        FilePathInfo.CreatedBy = UserID
                    Else
                        FilePathInfo.ModifiedBy = UserID
                    End If
                    Me.Save(moduleID, staffID, returnVal)
                End If
            Next
            'Adam Nall - Added to change the pProfile isDirty. It changes this on the save 
            '            of the single profileinfo but not the parent class isDirty
            'J. Cockrum - Interesting!  Me.colIsDirty has no code in the SET.  Voodoo?
            '
            Me.colIsDirty = False
            RaiseEvent FilePathIsDirty(Me.colIsDirty)
        End Sub
        Public Function GetNext() As String
            Return GetNextPrev(1)
        End Function
        Public Function GetPrevious() As String
            Return GetNextPrev(-1)
        End Function
        Private Function GetNextPrev(ByVal direction As Integer) As String
            Dim strArr() As String = colFilePaths.GetKeys()
            'Dim nArr(strArr.GetUpperBound(0)) As Integer
            'Dim y As String
            'For Each y In strArr
            '    nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            'Next
            'nArr.Sort(nArr)
            colIndex = Array.BinarySearch(strArr, Me.ID.ToString)
            If colIndex + direction > -1 And _
                colIndex + direction <= strArr.GetUpperBound(0) Then
                Return FilePaths.Item(strArr.GetValue(colIndex + direction)).ID.ToString
            Else
                Return FilePaths.Item(strArr.GetValue(colIndex)).ID.ToString
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oFilePathInfo = New MUSTER.Info.ProfileInfo
            'colFilePaths.Clear()
        End Sub
        Public Sub Reset()
            Dim oFilePathLocal As MUSTER.Info.ProfileInfo
            For Each oFilePathLocal In FilePaths.Values
                oFilePathLocal.Reset()
                oFilePathLocal.IsDirty = False
            Next
        End Sub
#End Region
#Region "Miscellaneous Operations"

        Public Function FilePathTable() As DataTable
            Try
                Dim dtFilePaths As New DataTable
                Dim oFilePathInfoLocal As MUSTER.Info.ProfileInfo
                Dim drRow As DataRow
                dtFilePaths.Columns.Add("FileName")
                dtFilePaths.Columns.Add("FilePath")
                Retrieve("SYSTEM|COMMON_PATHS", False)
                For Each oFilePathInfoLocal In Me.colFilePaths.Values
                    drRow = dtFilePaths.NewRow
                    drRow("FileName") = oFilePathInfoLocal.ProfileMod1
                    drRow("FilePath") = oFilePathInfoLocal.ProfileValue
                    dtFilePaths.Rows.Add(drRow)
                Next
                Return dtFilePaths
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
#End Region
#End Region
#Region "Events"

        Private Sub InfoBecameDirty(ByVal DirtyState As Boolean) Handles oFilePathInfo.InfoBecameDirty

            RaiseEvent FilePathIsDirty(Me.colIsDirty)
        End Sub

#End Region
    End Class
End Namespace
