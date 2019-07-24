'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.AppFlag
'   Provides the operations required to manipulate an Entity object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0          AN                   Original class definition
'  1.1          EN      02/22/2005   Added GetGUIDForRef function  and Modified GetValuePair,Retrieve function and Added New Attribute ModuleName... 
'  1.2          EN      02/22/2005   Modified Remove method to remove all the Guids related to that owner.. 
'  1.3          EN      02/24/2005   Added Ucase$ in checking 

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
' NOTE: This file to be used as AppFlag to build other objects.
'       Replace keyword "AppFlag" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pAppFlag
#Region "Public Events"
        Public Event AppFlagErr(ByVal MsgStr As String)
        Public Event AppFlagChanged(ByVal bolValue As Boolean)
        Public Event ColChanged(ByVal bolValue As Boolean)

        Public Event ActiveFormChanged(ByVal FormGUID As String, ByRef FormInst As Windows.Forms.Form)
        Public Event EntityChanged(ByVal FormGUID As String, ByVal EntityKey As String, ByRef FormInst As Windows.Forms.Form)
        Public Event ActivateWindow(ByVal WindowTitle As String, ByRef MyForm As Windows.Forms.Form)
        Public Event ActivateMusterControls(ByVal WindowTitle As String, ByRef MyForm As Windows.Forms.Form)
#End Region
#Region "Private Member Variables"
        Private oAppFlagInfo As Muster.Info.AppFlagInfo
        Private WithEvents colAppFlags As Muster.Info.AppFlagsCollection
        'Private oAppFlagDB As New Muster.DataAccess.AppFlagDB
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private nID As Int64 = -1
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private refForm As Windows.Forms.Form
#End Region
#Region "Constructors"
        Public Sub New()
            oAppFlagInfo = New Muster.Info.AppFlagInfo
            colAppFlags = New Muster.Info.AppFlagsCollection
        End Sub

        Public Sub New(ByVal refForm As Windows.Forms.Form)
            oAppFlagInfo = New Muster.Info.AppFlagInfo
            colAppFlags = New Muster.Info.AppFlagsCollection
            Me.refForm = refForm
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property Key() As String
            Get
                Return oAppFlagInfo.Key
            End Get
            Set(ByVal Value As String)
                oAppFlagInfo.Key = Value
            End Set
        End Property
        Public Property Value() As Object
            Get
                Return oAppFlagInfo.Value
            End Get
            Set(ByVal Value As Object)
                oAppFlagInfo.Value = Value
            End Set
        End Property
        Public Property ModuleName() As String
            Get
                Return oAppFlagInfo.ModuleName
            End Get
            Set(ByVal Value As String)
                oAppFlagInfo.ModuleName = Value
            End Set
        End Property

#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        'Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal ref As String, ByVal key As String, ByVal val As Object, Optional ByVal strModule As String = "ModuleName") As MUSTER.Info.AppFlagInfo
            Dim oAppFlagInfoLocal As MUSTER.Info.AppFlagInfo
            Try
                For Each oAppFlagInfoLocal In colAppFlags.Values
                    If oAppFlagInfoLocal.Key = ref + " - " + Trim(UCase$(key)) Then
                        oAppFlagInfo = oAppFlagInfoLocal
                        colAppFlags.Remove(oAppFlagInfo)
                        Exit For
                    End If
                Next

                oAppFlagInfoLocal = New MUSTER.Info.AppFlagInfo
                oAppFlagInfoLocal.Key = ref + " - " + Trim(UCase$(key))
                oAppFlagInfoLocal.Value = val
                If UCase$(strModule) <> UCase$("ModuleName") Then
                    oAppFlagInfoLocal.ModuleName = Trim(UCase$(strModule))
                End If
                oAppFlagInfo = oAppFlagInfoLocal
                colAppFlags.Add(oAppFlagInfo)

                If ref + " - " + key = "0 - ActiveForm" Then
                    Dim ActiveWindowName As String = Me.GetValuePair(val.ToString, "WindowName")
                    RaiseEvent ActiveFormChanged(val.ToString, refForm)
                    MakeWindowActive(ActiveWindowName)
                    If ActiveWindowName.StartsWith("Registration") Or _
                       ActiveWindowName.StartsWith("Technical") Or _
                       ActiveWindowName.StartsWith("Financial") Or _
                       ActiveWindowName.StartsWith("Closure") Or _
                       ActiveWindowName.StartsWith("Fees") Or _
                       ActiveWindowName.StartsWith("C & E") Then

                        RaiseEvent EntityChanged(val.ToString, "OwnerAddress", refForm)
                        RaiseEvent EntityChanged(val.ToString, "FacilityAddress", refForm)
                    End If
                    If ActiveWindowName.StartsWith("Company") Then
                        RaiseEvent EntityChanged(val.ToString, "CompanyID", refForm)
                    End If
                End If
                If key = "OwnerAddress" Or key = "FacilityAddress" Then
                    RaiseEvent EntityChanged(ref, key, refForm)
                End If
                If key = "CompanyID" Then
                    RaiseEvent EntityChanged(ref, key, refForm)
                End If
                Return oAppFlagInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try


        End Function

        Public Function Retrieve(ByVal AppFlagName As String) As MUSTER.Info.AppFlagInfo
            Try
                oAppFlagInfo = Nothing
                If colAppFlags.Contains(AppFlagName) Then
                    oAppFlagInfo = colAppFlags(AppFlagName)
                Else
                    If oAppFlagInfo Is Nothing Then
                        oAppFlagInfo = New MUSTER.Info.AppFlagInfo
                    End If
                    'oAppFlagInfo = oAppFlagDB.DBGetByName(AppFlagName)
                    'colAppFlags.Add(oAppFlagInfo)
                End If
                Return oAppFlagInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function


        Public Function GetValuePair(ByVal ref As String, ByVal key As String, Optional ByVal SupressErrorMsg As Boolean = False) As Object
            'If Not Me.Retrieve(ref, key, Nothing) Is Nothing Then
            '    Return Me.Retrieve(ref, key, Nothing).Value
            'End If

            Try
                Dim oAppFlagInfoLocal As New MUSTER.Info.AppFlagInfo
                oAppFlagInfoLocal = colAppFlags.Item(ref + " - " + Trim(UCase$(key)))
                If Not oAppFlagInfoLocal Is Nothing Then
                    Return oAppFlagInfoLocal.Value
                Else
                    Return Nothing
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function GetGUIDForRef(ByVal RefKey As String, ByVal RefValue As String, ByVal RefModule As String) As String

            Dim oAppFlagInfoLocal As New MUSTER.Info.AppFlagInfo
            GetGUIDForRef = String.Empty
            For Each oAppFlagInfoLocal In colAppFlags.Values
                If oAppFlagInfoLocal.Key.Length >= RefKey.Length Then
                    If oAppFlagInfoLocal.Key.EndsWith(Ucase$(RefKey.TrimEnd)) Then
                        If oAppFlagInfoLocal.Value = RefValue And UCase$(oAppFlagInfoLocal.ModuleName) = Trim(UCase$(RefModule)) Then
                            GetGUIDForRef = oAppFlagInfoLocal.Key.Remove(oAppFlagInfoLocal.Key.IndexOf(" - "), RefKey.TrimEnd.Length + 3)
                            Exit For
                        End If
                    End If
                End If
            Next
            Return GetGUIDForRef
        End Function


#End Region
#Region "Collection Operations"
        'Adds an entity to the collection as called for by ID
        'Public Sub Add(ByVal ID As Integer)
        '    Try
        '        oAppFlagInfo = New Muster.Info.AppFlagInfo
        '        oAppFlagInfo.Key = ref + " - " + Key
        '        oAppFlagInfo.Value = Value
        '        colAppFlags.Add(oAppFlagInfo)
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Sub
        ''Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oAppFlag As Muster.Info.AppFlagInfo)
            Try
                oAppFlagInfo = oAppFlag
                colAppFlags.Add(oAppFlagInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        'commented By Elango on Feb 22 2005 
        ''Public Sub Remove(ByVal ID As Integer)
        ''    Dim myIndex As Int16 = 1
        ''    Dim oAppFlagInfoLocal As Muster.Info.AppFlagInfo

        ''    Try
        ''        For Each oAppFlagInfoLocal In colAppFlags.Values
        ''            If oAppFlagInfoLocal.Key = ID Then
        ''                colAppFlags.Remove(oAppFlagInfoLocal)
        ''                Exit Sub
        ''            End If
        ''            myIndex += 1
        ''        Next
        ''    Catch Ex As Exception
        ''        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        ''        Throw Ex
        ''    End Try
        ''    Throw New Exception("Pipe " & ID.ToString & " is not in the collection of Pipes.")
        ''End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oAppFlag As Muster.Info.AppFlagInfo)
            Try
                colAppFlags.Remove(oAppFlag)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("AppFlag " & oAppFlag.Key & " is not in the collection of AppFlags.")
        End Sub

        Public Function Remove(ByVal ref As String, ByVal key As String) As Boolean
            Try
                colAppFlags.Remove(ref + " - " + key)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function Remove(ByVal GUIDKey As String) As Boolean
            Dim oAppFlagInfoLocal As MUSTER.Info.AppFlagInfo

            Try
                'SS
                'Dim DeleteList() As String
                Dim DeleteList() As String = New String() {}
                Dim cnt As Int16 = 0

                For Each oAppFlagInfoLocal In colAppFlags.Values
                    If oAppFlagInfoLocal.Key.Length >= GUIDKey.Length Then
                        If oAppFlagInfoLocal.Key.Substring(0, GUIDKey.Length) = GUIDKey.TrimEnd Then
                            ReDim Preserve DeleteList(cnt)
                            DeleteList(cnt) = oAppFlagInfoLocal.Key
                            cnt = cnt + 1
                        End If
                    End If
                Next
                If Not DeleteList Is Nothing Then
                    For cnt = 0 To UBound(DeleteList)
                        colAppFlags.Remove(DeleteList(cnt))
                    Next
                End If
                Return True
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetNext() As String
            Return GetNextPrev(1)
        End Function
        Public Function GetPrevious() As String
            Return GetNextPrev(-1)
        End Function
        Private Function GetNextPrev(ByVal direction As Integer) As String
            Dim strArr() As String = colAppFlags.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.Key.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colAppFlags.Item(nArr.GetValue(colIndex + direction)).Key.ToString
            Else
                Return colAppFlags.Item(nArr.GetValue(colIndex)).Key.ToString
            End If
        End Function
#End Region
#End Region
#Region "General Operations"
        Public Sub Clear()
            oAppFlagInfo = New MUSTER.Info.AppFlagInfo
        End Sub
        Public Sub Reset()
            oAppFlagInfo.Reset()
        End Sub

        Public Sub MakeWindowActive(ByVal ThisTitle As String)
            RaiseEvent ActivateWindow(ThisTitle, refForm)
        End Sub

        Public Sub ActivateAuxControls(ByVal ThisTitle As String)
            RaiseEvent ActivateMusterControls(ThisTitle, refForm)
        End Sub
#End Region

    End Class
End Namespace
