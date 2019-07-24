
Option Strict On
Option Explicit On 

Imports Infragistics.Win.UltraWinGrid
Imports System.Collections


Public Class GridMaster

#Region "EntityOpenClass"

    Public Class EntityItemClass

        Public moduleID As String
        Public keyword As String
        Public value As String
        Public tabPage As String
        Public activeControl As Control



        Sub New(ByVal modID As String, ByVal keywordStr As String, ByVal valueStr As String, ByVal tabPageStr As String, ByVal activeControlItem As Control)

            moduleID = modID
            keyword = keywordStr
            value = valueStr
            tabPage = tabPageStr
            activeControl = activeControlItem

        End Sub

    End Class

#End Region

#Region "Bot Class Structure"
    Public Class BotClass

        Public ID As Integer
        Public ModuleType As Type
        Public Command As String = String.Empty
        Public ModuleFormInProcess As Form = Nothing
        Public ObjectInProcess As Object
        Public DataFieldName As String


        Sub New(ByVal idNum As Integer, ByVal moduleFormType As Type, ByVal commandText As String, ByVal fieldname As String)
            ID = idNum
            ModuleType = moduleFormType
            Command = commandText
            DataFieldName = fieldname

        End Sub


    End Class
#End Region


#Region "Instance Construct"

    Protected Sub New()
        'initialization code goes here
    End Sub

#End Region


#Region "Shared Constants for commands"

    Public Const OPENCNERECORDBYID As String = "OpenCNERecordByID"

#End Region

#Region "Private Shared (Persisted) members"




    Private Shared _thisInstance As GridMaster
    Private Shared _padLock As New Object
#End Region

#Region "private members"


    Private _gridSorts As IList
    Private _gridSortDict As IDictionary
    Private _ignoreSortEvent As Boolean = False
    Private _bot As BotClass
    Private _firstWindow As Boolean = True
    Private _entityDict As IDictionary
    Private _entityList As IList
    Private _container As MusterContainer
    Private _ActivateClickedRow As Infragistics.Win.UltraWinGrid.UltraGridRow







#End Region


#Region "Public Properties (live)"

    Public ReadOnly Property Bot() As BotClass
        Get
            Return _bot
        End Get
    End Property


    Public Property GridDict() As IDictionary
        Get

            If _gridSortDict Is Nothing Then

                _gridSortDict = New System.Collections.SortedList

            End If

            Return _gridSortDict

        End Get

        Set(ByVal Value As IDictionary)
            _gridSortDict = Value
        End Set

    End Property


    Public Property EntityDict() As IDictionary
        Get

            If _entityDict Is Nothing Then

                _entityDict = New System.Collections.SortedList

            End If

            Return _entityDict

        End Get

        Set(ByVal Value As IDictionary)
            _entityDict = Value
        End Set

    End Property


    Public Property gridSorts() As IList

        Get

            If _gridSorts Is Nothing Then

                _gridSorts = New System.Collections.ArrayList

            End If

            Return _gridSorts

        End Get

        Set(ByVal Value As IList)
            _gridSorts = Value
        End Set

    End Property

    Public Property EntityList() As IList

        Get

            If _entityList Is Nothing Then
                _entityList = New System.Collections.ArrayList
            End If

            Return _entityList

        End Get

        Set(ByVal Value As IList)
            _entityList = Value
        End Set

    End Property

#End Region

#Region "Private Bot Search and Initiate Job"

    Private Function BotFoundAndInitiated(ByVal frm As Form, Optional ByVal obj As Object = Nothing) As Boolean

        If _bot.ModuleType.ToString = frm.GetType.ToString Then

            _bot.ModuleFormInProcess = frm

            _bot.ObjectInProcess = obj

            Return PerformBotJob()

        End If

        Return False

    End Function

#End Region

#Region "Instance Functions"

    Public Sub ActivateBot(ByVal recordID As Integer, ByVal formTypeToBeUsed As Type, ByVal botCommand As String, ByVal FieldName As String)
        _bot = New BotClass(recordID, formTypeToBeUsed, botCommand, FieldName)
    End Sub

    Public Sub DeActivateBot()
        _bot = Nothing
    End Sub


    Public Sub ReadyForBotJobs(ByVal frm As Form, Optional ByRef obj As Object = Nothing)

        If Not _bot Is Nothing Then

            If BotFoundAndInitiated(frm, obj) Then

                obj = _bot.ObjectInProcess

                DeActivateBot()

            End If

        End If

    End Sub



#End Region

#Region "public shared Function"


    Public Shared Function GlobalInstance() As GridMaster

        SyncLock _padLock

            If _thisInstance Is Nothing Then
                _thisInstance = New GridMaster
            End If

            Return _thisInstance

        End SyncLock

    End Function


#End Region


#Region "Live events"

    Sub GetGridsNotDevelopedDuringFormActivation(ByVal sender As Object, ByVal e As EventArgs)

        If DirectCast(sender, Control).Visible = True Then

            AttachFormToInstance(DirectCast(sender, Control))

        End If

    End Sub

    Sub UpdateRowSelection(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs)
        If e.Type.FullName.ToUpper.IndexOf("ULTRAGRIDROW") > -1 Then
            Try
                DirectCast(sender, Infragistics.Win.UltraWinGrid.UltraGrid).ActiveRow = DirectCast(sender, Infragistics.Win.UltraWinGrid.UltraGrid).Selected.Rows(0)
            Catch ex As Exception
                'swallow error if no row selection after all
            End Try
        End If


    End Sub



    Sub Columns_Sorted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BandEventArgs)

        If Not _ignoreSortEvent Then

            Dim name As String = DirectCast(sender, UltraGrid).Name
            Dim Band As String = e.Band.ToString

            If Not e.Band.SortedColumns Is Nothing AndAlso e.Band.SortedColumns.Count > 0 Then

                Dim Column As String = e.Band.SortedColumns(0).ToString
                Dim desc As String = CStr(IIf(e.Band.SortedColumns(0).SortIndicator = SortIndicator.Descending, "TRUE", "FALSE"))

                InsertGridSort(String.Format("{0};{1};{2};{3}", name, Band, Column, desc), String.Format("{0};{1}", name, Band))
            End If

        End If

    End Sub


    Sub refresh_Sort(ByVal sender As Object, ByVal e As EventArgs)

        Dim _controlList As IList = seekControlType(DirectCast(sender, Control), "Infragistics.Win.UltraWinGrid.UltraGrid")

        If Not _controlList Is Nothing AndAlso _controlList.Count > 0 Then

            For Each cntr As Control In _controlList

                ActivateLastUserSort(DirectCast(cntr, UltraGrid))

            Next
        End If

        _controlList.Clear()
        _controlList = Nothing



    End Sub

    Sub AutoSetTabPage(ByVal sender As Object, ByVal e As EventArgs)

        If _container.GoToTabPage <> String.Empty Then
            With DirectCast(sender, TabControl)

                For Each tab As TabPage In .TabPages
                    If tab.Name.ToUpper = _container.GoToTabPage Then
                        .SelectedTab = tab
                        Exit For
                    End If

                Next

            End With
        End If



    End Sub

    Sub LoadEntityDictionary(ByVal sender As Object, ByVal e As EventArgs)

        Dim ownerKey As String = String.Empty
        Dim facKey As String = String.Empty
        Dim frmKey As String = String.Empty

        Dim key As String = String.Empty


        If Not _container Is Nothing AndAlso Not _container.ActiveMdiChild Is Nothing AndAlso _container.ActiveMdiChild.Text.ToUpper.IndexOf("SEARCH RESULT") = -1 AndAlso Not _container.pOwn Is Nothing AndAlso _container.pOwn.ID > 0 AndAlso _container.AppSemaphores.ModuleName <> String.Empty Then


            With _container

                frmKey = .ActiveMdiChild.GetType.ToString

                frmKey = frmKey.Substring(frmKey.LastIndexOf(".") + 1).Replace("and", " & ")

                If frmKey = "Inspection" Then
                    frmKey = "inspectionSchedule"
                End If

                If Not .pOwn.Facility Is Nothing Then
                    facKey = .pOwn.Facility.ID.ToString
                Else
                    facKey = "-1"
                End If

                If DirectCast(sender, Control).Name.ToUpper.IndexOf("OWNERDETAIL") <> -1 OrElse (TypeOf sender Is TabControl AndAlso DirectCast(sender, TabControl).SelectedTab.Name.ToUpper.IndexOf("OWNER") <> -1) Then
                    facKey = "-1"
                End If

                ownerKey = .pOwn.ID.ToString

            End With




            key = String.Format("{0}", _container.ActiveMdiChild.Text.Replace("  ", " ").Trim)


            If Not EntityDict.Contains(key) Then

                Dim newentity As EntityItemClass
                Dim cnt As Integer = 0

                If TypeOf sender Is TabControl Then
                    newentity = New EntityItemClass(frmKey, IIf(facKey <> "-1", "Facility ID", "Owner ID").ToString, IIf(facKey <> "-1", facKey, ownerKey).ToString, DirectCast(sender, TabControl).SelectedTab.Name.ToUpper, _container.ActiveMdiChild.ActiveControl)
                Else
                    newentity = New EntityItemClass(frmKey, IIf(facKey <> "-1", "Facility ID", "Owner ID").ToString, IIf(facKey <> "-1", facKey, ownerKey).ToString, String.Empty, _container.ActiveMdiChild.ActiveControl)
                End If

                EntityDict.Add(key, newentity)
                EntityList.Add(key)

                DirectCast(EntityList, Collections.ArrayList).Sort()

                _container.MnHistoryItems.MenuItems.Clear()

                For Each item As String In EntityList
                    _container.MnHistoryItems.MenuItems.Add(item, New EventHandler(AddressOf openForm))
                Next

            End If

        Else

            Return

        End If
    End Sub

    Public Sub CleanEntityDictionary()
        EntityDict.Clear()
        EntityList.Clear()
    End Sub

    Sub openForm(ByVal sender As Object, ByVal e As EventArgs)

        With _container

            If EntityDict.Contains(DirectCast(sender, MenuItem).Text) Then

                Dim eItem As EntityItemClass = DirectCast(EntityDict.Item(DirectCast(sender, MenuItem).Text), EntityItemClass)

                .txtOwnerQSKeyword.Text = eItem.value
                .cmbSearchModule.Text = eItem.moduleID
                .cmbQuickSearchFilter.Text = eItem.keyword
                .InvokeQuickOwneerButtonClick()

                If eItem.tabPage <> String.Empty AndAlso Not _container.ActiveMdiChild Is Nothing Then

                    Dim tabcontrolList As IList = Me.seekControlType(_container.ActiveMdiChild, "System.Windows.Forms.TabControl")

                    If Not tabcontrolList Is Nothing Then
                        For Each item As TabControl In tabcontrolList

                            For Each page As TabPage In item.TabPages
                                If page.Name.ToUpper = eItem.tabPage Then
                                    item.SelectedTab = page

                                    If Not eItem.activeControl Is Nothing Then
                                        Dim objs As IList = Me.seekControlType(page, eItem.activeControl.GetType.ToString)

                                        If Not objs Is Nothing Then

                                            For Each obj As Control In objs

                                                If obj.Name.ToUpper = eItem.activeControl.Name Then
                                                    obj.Visible = True
                                                    obj.Focus()
                                                End If
                                            Next
                                        End If
                                    End If

                                End If
                            Next
                        Next
                    End If

                End If

            End If

        End With

    End Sub
#End Region

#Region "Bot Jobs"

    Private Function PerformBotJob() As Boolean


        Dim isJobDone As Boolean = True

        Try

            Select Case _bot.Command

                Case GridMaster.OPENCNERECORDBYID
                    OpenFacilityRecordInCNEForm()
                Case Else
                    isJobDone = False
            End Select

        Catch ex As Exception

            MsgBox(String.Format("{0}{1}{1}{2}", "Error Performing BOT JOB!", vbCrLf, ex.ToString), MsgBoxStyle.OKOnly, "Grid Master Bot Job Exception")
            isJobDone = False

        End Try

        Return isJobDone

    End Function


    Private Sub OpenFacilityRecordInCNEForm()

        With DirectCast(Me.Bot.ModuleFormInProcess, CandEManagement)

            Dim arr() As Integer = DirectCast(Bot.ObjectInProcess, Integer())

            If Bot.DataFieldName = "FACILITY_ID" Then
                arr(0) = 0
                arr(1) = Bot.ID
            Else
                arr(0) = Bot.ID
                arr(1) = 0
            End If


            Bot.ObjectInProcess = arr
        End With

    End Sub

#End Region


#Region "Live Instance Functions"


    Function seekControlType(ByVal cntrl As Control, ByVal thisType As String) As IList

        Dim obj As IList = Nothing


        For Each cntr As Control In cntrl.Controls

            If cntr.Controls.Count > 0 Then

                Dim obj2 As IList = seekControlType(cntr, thisType)

                If Not obj2 Is Nothing Then

                    If obj Is Nothing Then
                        obj = New System.Collections.ArrayList
                    End If

                    For Each objItem As Control In obj2
                        obj.Add(objItem)
                    Next

                    obj2.Clear()
                    obj2 = Nothing

                End If
            End If

            If cntr.GetType.ToString.ToUpper = thisType.ToUpper.Trim Then


                If obj Is Nothing Then
                    obj = New System.Collections.ArrayList
                End If

                obj.Add(cntr)
            End If


        Next

        Return obj



    End Function




    Sub AttachFormToInstance(ByVal thisForm As Control)

        Try

            Dim _controlList As IList = seekControlType(thisForm, "Infragistics.Win.UltraWinGrid.UltraGrid")


            If Not _controlList Is Nothing AndAlso _controlList.Count > 0 Then

                For Each item As UltraGrid In _controlList
                    AttachGridToInstance(item)
                Next
                _controlList.Clear()

            End If

            _controlList = Nothing
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()

        End Try



    End Sub



    Sub AttachGridToInstance(ByVal thisGrid As UltraGrid)

        RemoveHandler thisGrid.AfterSortChange, AddressOf Columns_Sorted
        AddHandler thisGrid.AfterSortChange, AddressOf Columns_Sorted

        RemoveHandler thisGrid.AfterSelectChange, AddressOf UpdateRowSelection
        AddHandler thisGrid.AfterSelectChange, AddressOf UpdateRowSelection


        Dim tabform As Control = thisGrid


        Try
            Do

                While Not TypeOf tabform Is TabControl AndAlso Not TypeOf tabform Is Form AndAlso Not tabform.Parent Is Nothing
                    tabform = tabform.Parent
                End While

                If TypeOf tabform Is TabPage Then

                    With DirectCast(tabform, TabPage)

                        RemoveHandler .VisibleChanged, AddressOf LoadEntityDictionary
                        AddHandler .VisibleChanged, AddressOf LoadEntityDictionary

                        RemoveHandler .Validated, AddressOf LoadEntityDictionary
                        AddHandler .Validated, AddressOf LoadEntityDictionary

                        LoadEntityDictionary(tabform, New EventArgs)

                    End With

                End If

                If TypeOf tabform Is TabControl Then

                    With DirectCast(tabform, TabControl)

                        RemoveHandler .SelectedIndexChanged, AddressOf refresh_Sort
                        AddHandler .SelectedIndexChanged, AddressOf refresh_Sort

                        RemoveHandler .SelectedIndexChanged, AddressOf LoadEntityDictionary
                        AddHandler .SelectedIndexChanged, AddressOf LoadEntityDictionary

                        RemoveHandler .VisibleChanged, AddressOf AutoSetTabPage
                        AddHandler .VisibleChanged, AddressOf AutoSetTabPage


                    End With

                    tabform = tabform.Parent

                End If

                If TypeOf tabform Is Form Then

                    With DirectCast(tabform, Form)

                        If _container Is Nothing Then
                            _container = DirectCast(DirectCast(tabform, Form).MdiParent, MusterContainer)
                        End If

                        RemoveHandler .CausesValidationChanged, AddressOf refresh_Sort

                        AddHandler .CausesValidationChanged, AddressOf refresh_Sort

                        LoadEntityDictionary(tabform, New EventArgs)

                    End With

                End If

            Loop Until TypeOf tabform Is Form

        Catch ex As Exception

            Throw ex

        End Try


    End Sub


    Sub InsertGridSort(ByVal value As String, ByVal key As String)

        If GridDict.Contains(key) Then

            gridSorts.Remove(GridDict.Item(key))
            GridDict.Remove(key)

        End If

        GridDict.Add(key, value)
        gridSorts.Add(value)
    End Sub

    Sub ActivateLastUserSort(ByVal curgrid As Infragistics.Win.UltraWinGrid.UltraGrid)



        If curgrid.Visible = True Then
            For Each item As String In gridSorts

                If item.Substring(0, item.IndexOf(";")) = curgrid.Name Then

                    _ignoreSortEvent = True


                    Threading.Thread.Sleep(100)
                    Dim band As String = item.Substring(item.IndexOf(";") + 1)
                    Dim column As String = band.Substring(band.IndexOf(";") + 1)
                    Dim desc As String = column.Substring(column.IndexOf(";") + 1)


                    If curgrid.DisplayLayout.Bands.Contains(band) AndAlso curgrid.DisplayLayout.Bands(band).Columns.Contains(column) Then

                        band = band.Substring(0, band.IndexOf(";"))

                        column = column.Substring(0, column.IndexOf(";"))

                        curgrid.DisplayLayout.Bands(band).SortedColumns.Clear()

                        curgrid.DisplayLayout.Bands(band).SortedColumns.Add(curgrid.DisplayLayout.Bands(band).Columns(column), (desc = "TRUE"), False)
                        curgrid.DisplayLayout.Bands(band).SortedColumns.RefreshSort(False)


                        _ignoreSortEvent = True
                    End If


                End If

            Next
        End If



    End Sub



#End Region







End Class
