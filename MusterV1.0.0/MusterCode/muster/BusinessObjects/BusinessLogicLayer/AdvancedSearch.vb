'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.Address
'   Provides the info and collection objects to the client for manipulating
'   an AddressInfo object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'   1.0     EN          12/10/2004   Original class definition.
'   1.1     EN          02/01/2005   Added all methods.
'   1.2     EN          02/02/2005   Added New Event and New method ChangeOrder
'   1.3     JVC2        02/02/2005   Changed column header Facility FipsCode to
'                                       Facility County in GetAdvSearchTable.
'-------------------------------------------------------------------------------
'
' TODO - Integrate with solution 2/3/05
'
Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pAdvancedSearch
        Inherits pFavSearch

#Region "Public Events"
        Public Event AdvancedSearchErr(ByVal MsgStr As String, ByVal strSrc As String)
        Public Event eEnableDisableDelete(ByVal ncount As Integer, ByVal strSrc As String)
#End Region
#Region "Private Member Variables"
        Private oAdvancedSearchDB As New Muster.DataAccess.AdvancedSearchDB
        Private strTankStatus As String
        Private dtCriteria As New DataTable
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Public Sub New()

        End Sub
#End Region
#Region "Exposed attributes"
        'Public Property TankStatus() As String
        '    Get
        '        Return strTankStatus
        '    End Get
        '    Set(ByVal Value As String)
        '        strTankStatus = Value
        '    End Set
        'End Property
        Public Property Search_Type() As String
            Get
                Return Me.SearchType
            End Get
            Set(ByVal Value As String)
                Me.SearchType = Value
            End Set
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        Public Sub AddNewInfo(ByVal nSEARCH_ID As Integer, ByVal strName As String, ByVal strSearchType As String)
            Try
                Dim strChildName As String
                Dim oChildLocal As New Muster.Info.FavSearchChildInfo
                Dim oChild As New Muster.Info.FavSearchChildInfo
                Dim nOrder As Integer = 1
                Dim nTempOrder As Integer
                Dim bolOrderIncrease As Boolean

                If Len(Trim(strSearchType)) = 0 Then
                    RaiseEvent AdvancedSearchErr("Select Search Type", Me.ToString)
                    Exit Sub
                End If
                If Len(strName) = 0 Then
                    RaiseEvent AdvancedSearchErr("Select an available SearchBy from the list", Me.ToString)
                    Exit Sub
                End If
                If nSEARCH_ID = 0 Then 'if 0 then new Search so we have to Add Parent First 
                    ' Dim oParent As New MUSTER.Info.FavSearchParentInfo
                    Me.AddParent(oParentInfo)
                    'Me.Retrieve(nSEARCH_ID)
                    If strName.IndexOf("Owner") = 0 Then
                        strChildName = "Owner Name"
                    ElseIf strName.IndexOf("Facility") = 0 Then
                        strChildName = "Facility Name"
                    ElseIf strName.IndexOf("Contact") = 0 Then
                        strChildName = "Contact Name"
                    ElseIf strName.IndexOf("Contractor") = 0 Then
                        strChildName = "Licensee Name"
                    ElseIf strName.IndexOf("Company") = 0 Then
                        strChildName = "Company Name"
                    ElseIf strName.IndexOf("All") = 0 Then
                        strChildName = "Name"
                    End If
                    'nOrder = 1
                Else
                    'Existing Parent and adding new child so check for duplicate entries.. 
                    strChildName = strName
                    Me.Retrieve(nSEARCH_ID)
                    bolOrderIncrease = False
                    For Each oChildLocal In Me.colChildren.Values
                        If oChildLocal.ParentID = oParentInfo.ID Then
                            If nOrder <= oChildLocal.Order Then
                                nOrder = oChildLocal.Order
                                bolOrderIncrease = True
                            End If
                            If oChildLocal.CriterionName = strChildName Then
                                RaiseEvent AdvancedSearchErr("Duplicate Entry. Please select a new item from the List", Me.ToString)
                                Exit Sub
                            End If
                        End If
                    Next
                    If bolOrderIncrease Then
                        nOrder += 1
                    End If
                End If
                Me.AddChild(oChild)
                oChild.CriterionName = strChildName
                oChild.Order = nOrder
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function GetCriteria(ByVal nSEARCH_ID As Integer) As DataTable
            Try
                dtCriteria = Nothing
                Me.Retrieve(nSEARCH_ID)
                dtCriteria = Me.ChildTable
                Return dtCriteria
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetAdvSearchTable(ByVal SearchType As String, ByVal dtInputTable As DataTable) As DataTable
            Dim ReturnTable As New DataTable
            Dim dr As DataRow
            Dim sno As Int64 = 1
            Dim i As Integer
            Dim drow As DataRow
            Dim dsSet As DataSet
            Try
                If SearchType = "Facility" Then
                    ReturnTable.Columns.Add("SNo", Type.GetType("System.Int64"))
                    ReturnTable.Columns.Add("Facility ID", Type.GetType("System.Int64"))
                    ReturnTable.Columns.Add("Owner Name", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("Facility Name", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("Facility Address", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("Facility City", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("Facility County", Type.GetType("System.String"))
                    'ReturnTable.Columns.Add("Facility Points", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("Points", Type.GetType("System.Int64"))
                    'ReturnTable.Columns.Add("Current Lust Site", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("Latitude", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("Longitude", Type.GetType("System.String"))

                    ReturnTable.Columns.Add("Lust Site", Type.GetType("System.String"))
                    'ReturnTable.Columns.Add("Currently In Use Tanks", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("CIU", Type.GetType("System.Int64"))
                    'ReturnTable.Columns.Add("Temporarily Out of Use", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("TOS", Type.GetType("System.Int64"))
                    'ReturnTable.Columns.Add("Permanently Out of Use", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("TOSI", Type.GetType("System.Int64"))
                    ReturnTable.Columns.Add("POU", Type.GetType("System.Int64"))
                    'ReturnTable.Columns.Add("Permanent Closure Pending", Type.GetType("System.Int64"))
                    'ReturnTable.Columns.Add("Temporarily Out of Use Indefinitely", Type.GetType("System.String"))
                    'ReturnTable.Columns.Add("TOSI", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("Total Tanks", Type.GetType("System.Int64"))

                    'ReturnTable.Columns.Add("Permanent Closure Pending", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("Facility Contact", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("Facility Phone", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("Owner ID", Type.GetType("System.Int64"))

                    For Each drow In dtInputTable.Rows
                        dr = ReturnTable.NewRow
                        dr.Item("SNo") = sno
                        dr.Item("Facility ID") = drow("Facility_ID")
                        dr.Item("Facility Name") = drow("Facility_Name")
                        dr.Item("Owner Name") = drow("Owner_Name")

                        dr.Item("Facility Address") = drow("Facility_Address")
                        dr.Item("Longitude") = drow("Longitude")
                        dr.Item("Latitude") = drow("Latitude")

                        dr.Item("Facility City") = drow("Facility_City")
                        dr.Item("Facility County") = drow("Facility_County")
                        dr.Item("Facility Phone") = drow("Facility_Phone")
                        ' #2915
                        If dtInputTable.Select("Facility_ID = " + drow("Facility_ID").ToString + " and (Current_Lust_Status = 'OPEN' or Current_Lust_Status = 'CLOSED')").Length > 0 Then
                            dr.Item("Lust Site") = "Yes/" + drow("Current_Lust_Status")
                        Else
                            dr.Item("Lust Site") = "No"
                        End If
                        'If drow("Current_Lust_Status") <> String.Empty Then
                        '    If UCase(drow("Current_Lust_Status")) = UCase("open") Then
                        '        dr.Item("Lust Site") = "Yes"
                        '    ElseIf UCase(drow("Current_Lust_Status")) = UCase("Closed") Then
                        '        'dr.Item("Lust Site") = drow("Current_Lust_Status")
                        '        dr.Item("Lust Site") = "No"
                        '    End If
                        'Else
                        '    dr.Item("Lust Site") = drow("Current_Lust_Status")
                        'End If
                        dr.Item("CIU") = drow("Currently_In_Use")
                        dr.Item("Total Tanks") = drow("Total_Tanks")
                        dr.Item("TOS") = drow("Temporarily_Out_of_Use")
                        dr.Item("POU") = drow("Permanently_Out_of_Use")
                        'dr.Item("Permanent Closure Pending") = drow("Permanent_Closure_Pending")
                        dr.Item("TOSI") = drow("Temporarily_Out_of_Use_Indefinitely")
                        dr.Item("Facility Contact") = drow("Facility_Contact")
                        dr.Item("Points") = drow("Points")
                        dr.Item("Owner ID") = drow("Owner_ID")
                        ReturnTable.Rows.Add(dr)
                        sno += 1
                    Next
                ElseIf SearchType = "Owner" Then
                    ReturnTable.Columns.Add("SNo", Type.GetType("System.Int64"))
                    ReturnTable.Columns.Add("Owner ID", Type.GetType("System.Int64"))
                    ReturnTable.Columns.Add("Owner Name", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("Owner Address", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("Owner City", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("State", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("Owner Contact", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("Owner Phone", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("Owner Points", Type.GetType("System.Int64"))
                    For Each drow In dtInputTable.Rows
                        dr = ReturnTable.NewRow
                        dr.Item("SNo") = sno
                        dr.Item("Owner ID") = drow("Owner_ID")
                        dr.Item("Owner Name") = drow("Owner_Name")
                        dr.Item("Owner Address") = drow("Address")
                        dr.Item("Owner City") = drow("CITY")
                        dr.Item("State") = drow("STATE")
                        dr.Item("Owner Contact") = drow("Contact")
                        dr.Item("Owner Phone") = drow("Phone")
                        dr.Item("Owner Points") = drow("POINTS")
                        ReturnTable.Rows.Add(dr)
                        sno += 1
                    Next
                ElseIf SearchType = "Contact" Then
                    ReturnTable.Columns.Add("Contact Name", Type.GetType("System.String"))
                    'ReturnTable.Columns.Add("Contact Last Name", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("Contact Address", Type.GetType("System.String"))
                    'ReturnTable.Columns.Add("Contact City", Type.GetType("System.String"))
                    'ReturnTable.Columns.Add("Contact State", Type.GetType("System.String"))
                    'ReturnTable.Columns.Add("Contact Phone", Type.GetType("System.String"))
                    'ReturnTable.Columns.Add("ZipCode", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("Contact Type", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("Associated At", Type.GetType("System.String"))

                    ReturnTable.Columns.Add("Contact Source", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("Contact Source ID", Type.GetType("System.Int64"))
                    ReturnTable.Columns.Add("Contact Points", Type.GetType("System.Int64"))
                    ReturnTable.Columns.Add("Module ID", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("Owner ID", Type.GetType("System.Int64"))
                    ReturnTable.Columns.Add("Facility ID", Type.GetType("System.Int64"))

                    For Each drow In dtInputTable.Rows
                        dr = ReturnTable.NewRow
                        dr.Item("Contact Name") = drow("Contact_Name")
                        'dr.Item("Contact Last Name") = drow("Contact_Last_Name")
                        dr.Item("Contact Address") = String.Format("{0}  {1}{5}{2}   {3}   {4}", drow("Address"), drow("City"), drow("State"), _
                                                                   drow("ZipCode"), drow("Phone_Number_One"), IIf(drow("City") Is DBNull.Value OrElse drow("City") = "", "", ", ")).Trim
                        'dr.Item("Contact City") = drow("City")
                        'dr.Item("Contact State") = drow("State")
                        'dr.Item("ZipCode") = drow("ZipCode")
                        dr.Item("Contact Type") = drow("Type")
                        dr.Item("Associated At") = drow("Associated At")
                        'dr.Item("Contact Phone") = drow("phone_number_one")
                        dr.Item("Contact Source") = drow("Contact Source")
                        dr.Item("Contact Source ID") = drow("Contact Source ID")
                        dr.Item("Contact Points") = drow("Points")
                        dr.Item("Module ID") = drow("ModuleID")
                        If dr.Item("Contact Source") = "Lust Event" Then
                            dsSet = Nothing
                            dsSet = oAdvancedSearchDB.DBGetEntity("SELECT * FROM V_LUSTOWNER WHERE EVENT_ID=" + dr.Item("Contact Source ID").ToString)

                            If Not dsSet Is Nothing Then
                                If dsSet.Tables(0).Rows.Count > 0 Then
                                    dr.Item("Owner ID") = dsSet.Tables(0).Rows(0).Item("Owner_ID")
                                    dr.Item("Facility ID") = dsSet.Tables(0).Rows(0).Item("Facility_ID")
                                Else
                                    dr.Item("Owner ID") = drow("Contact Source ID")
                                    dr.Item("Facility ID") = drow("Contact Source ID")
                                End If
                            Else
                                dr.Item("Owner ID") = drow("Contact Source ID")
                                dr.Item("Facility ID") = drow("Contact Source ID")
                            End If
                        ElseIf dr.Item("Contact Source") = "FinancialEvent" Then
                            dsSet = Nothing
                            dsSet = oAdvancedSearchDB.DBGetEntity("SELECT * FROM V_FINOWNER WHERE FIN_EVENT_ID=" + dr.Item("Contact Source ID").ToString)

                            If Not dsSet Is Nothing Then
                                If dsSet.Tables(0).Rows.Count > 0 Then
                                    dr.Item("Owner ID") = dsSet.Tables(0).Rows(0).Item("Owner_ID")
                                    dr.Item("Facility ID") = dsSet.Tables(0).Rows(0).Item("Facility_ID")
                                Else
                                    dr.Item("Owner ID") = drow("Contact Source ID")
                                    dr.Item("Facility ID") = drow("Contact Source ID")
                                End If
                            Else
                                dr.Item("Owner ID") = drow("Contact Source ID")
                                dr.Item("Facility ID") = drow("Contact Source ID")
                            End If
                        ElseIf dr.Item("Contact Source") = "Facility" Then
                            dsSet = Nothing
                            dsSet = oAdvancedSearchDB.DBGetEntity("SELECT Owner_id FROM tblreg_Facility where facility_ID=" + dr.Item("Contact Source ID").ToString)

                            If Not dsSet Is Nothing Then
                                If dsSet.Tables(0).Rows.Count > 0 Then
                                    dr.Item("Owner ID") = dsSet.Tables(0).Rows(0).Item("Owner_ID")
                                Else
                                    dr.Item("Owner ID") = drow("Contact Source ID")
                                End If
                            Else
                                dr.Item("Owner ID") = drow("Contact Source ID")
                            End If
                        ElseIf dr.Item("Contact Source") = "Closure Event" Then
                            dsSet = Nothing
                            dsSet = oAdvancedSearchDB.DBGetEntity("SELECT * FROM V_CLOSUREOWNER WHERE CLOSURE_ID=" + dr.Item("Contact Source ID").ToString)

                            If Not dsSet Is Nothing Then
                                If dsSet.Tables(0).Rows.Count > 0 Then
                                    dr.Item("Owner ID") = dsSet.Tables(0).Rows(0).Item("Owner_ID")
                                    dr.Item("Facility ID") = dsSet.Tables(0).Rows(0).Item("Facility_ID")
                                Else
                                    dr.Item("Owner ID") = drow("Contact Source ID")
                                    dr.Item("Facility ID") = drow("Contact Source ID")
                                End If
                            Else
                                dr.Item("Owner ID") = drow("Contact Source ID")
                                dr.Item("Facility ID") = drow("Contact Source ID")
                            End If
                        End If

                        ReturnTable.Rows.Add(dr)
                        sno += 1
                    Next
                ElseIf SearchType = "Company" Then
                    ReturnTable.Columns.Add("Company Name", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("Company Address", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("City", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("State", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("Zip", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("Phone", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("Installer/Closures", Type.GetType("System.Boolean"))
                    ReturnTable.Columns.Add("Closures", Type.GetType("System.Boolean"))
                    ReturnTable.Columns.Add("ERAC", Type.GetType("System.Boolean"))
                    ReturnTable.Columns.Add("Points", Type.GetType("System.Int64"))
                    ReturnTable.Columns.Add("Company ID", Type.GetType("System.Int64"))

                    For Each drow In dtInputTable.Rows
                        dr = ReturnTable.NewRow
                        dr.Item("Company Name") = drow("Company_Name")
                        dr.Item("Company Address") = drow("Address")
                        dr.Item("City") = drow("City")
                        dr.Item("State") = drow("State")
                        dr.Item("Zip") = drow("Zip")
                        dr.Item("Phone") = drow("phone_number_one")
                        dr.Item("Installer/Closures") = drow("CTIAC")
                        dr.Item("Closures") = drow("CTC")
                        dr.Item("ERAC") = drow("ERAC")
                        dr.Item("Points") = drow("Points")
                        dr.Item("Company ID") = drow("Company_ID")
                        ReturnTable.Rows.Add(dr)
                        sno += 1
                    Next
                ElseIf SearchType = "Contractor" Then
                    ReturnTable.Columns.Add("Licensee Name", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("Company Name", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("Address", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("City", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("State", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("Zip", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("Phone", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("Status", Type.GetType("System.String"))
                    ReturnTable.Columns.Add("Expiration Date", Type.GetType("System.DateTime"))
                    ReturnTable.Columns.Add("Points", Type.GetType("System.Int64"))
                    ReturnTable.Columns.Add("Licensee ID", Type.GetType("System.Int64"))
                    ReturnTable.Columns.Add("Company ID", Type.GetType("System.Int64"))

                    For Each drow In dtInputTable.Rows
                        dr = ReturnTable.NewRow
                        dr.Item("Licensee Name") = drow("Licensee Name")
                        dr.Item("Company Name") = drow("Company Name")
                        dr.Item("Address") = drow("Address")
                        dr.Item("City") = drow("City")
                        dr.Item("State") = drow("State")
                        dr.Item("Zip") = drow("Zip")
                        dr.Item("Phone") = drow("phone")
                        dr.Item("Status") = drow("Status")
                        dr.Item("Expiration Date") = drow("Expiration Date")
                        dr.Item("Points") = drow("Points")
                        dr.Item("Licensee ID") = drow("LicenseeID")
                        dr.Item("Company ID") = drow("Company_id")
                        ReturnTable.Rows.Add(dr)
                        sno += 1
                    Next
                End If
                Return ReturnTable
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try

        End Function
        Public Function GetResults(ByVal nSEARCH_ID As Integer, ByVal strSearchType As String, ByVal strTankstatus As String, ByVal nLustStatus As Integer) As DataSet
            Try
                Dim StrCriteriaInput As String
                Dim drow As DataRow
                Dim dtReturn As DataSet
                Dim drCriteria() As DataRow

                If Len(Trim(strSearchType)) = 0 Then
                    RaiseEvent AdvancedSearchErr("Select Search Type", Me.ToString)
                    Exit Function
                End If
                'Call GetCriteria if user has changed the collection information  or for new search ID .we need to get the new datatable.
                If Me.colIsDirty Then 'Checking Children Collection is dirty 
                    dtCriteria = GetCriteria(nSEARCH_ID)
                Else
                    ' use the same datatable dtcriteria
                End If
                If dtCriteria.Rows.Count = 0 Then
                    RaiseEvent AdvancedSearchErr("Enter Search Bys / Look Fors Filters", Me.ToString)
                    Exit Function
                End If
                For Each drow In dtCriteria.Rows
                    If IsDBNull(drow("CRITERION_VALUE")) Then
                        RaiseEvent AdvancedSearchErr("LookFors value should not be Empty for selected SearchBys.", Me.ToString)
                        Exit Function
                    End If
                    If drow("CRITERION_VALUE") = "" Then
                        RaiseEvent AdvancedSearchErr("LookFors value should not be Empty for selected SearchBys.", Me.ToString)
                        Exit Function
                    End If
                    If drow("CRITERION_NAME") = "Owner ID" Or drow("CRITERION_NAME") = "Facility ID" Or drow("CRITERION_NAME") = "Facility AIID" Then
                        If IsNumeric(drow("CRITERION_VALUE")) = False Then
                            RaiseEvent AdvancedSearchErr("Enter Valid Integer LookFors for selected SearchBys: " + drow("CRITERION_NAME").ToString, Me.ToString)
                            Exit Function
                        End If
                    End If
                    'StrCriteriaInput += IIf(StrCriteriaInput = String.Empty, "", ";") + drow("CRITERION_NAME") + ";" + drow("CRITERION_VALUE")
                Next

                drCriteria = dtCriteria.Select("", "CRITERION_ORDER ASC")

                Dim strCriterianName As String = String.Empty
                Dim strCriterianValue As String = String.Empty

                For i As Integer = 0 To UBound(drCriteria)
                    strCriterianName = drCriteria(i).Item("CRITERION_NAME")
                    strCriterianValue = drCriteria(i).Item("CRITERION_VALUE")

                    Select Case strCriterianName
                        Case "Name", "Address", "City"
                            If strSearchType = "Contractor" Then
                                strCriterianName = Trim("Licensee") + " " + Trim(strCriterianName)
                            Else
                                strCriterianName = Trim(strSearchType) + " " + Trim(strCriterianName)
                            End If
                        Case "Phone"
                            strCriterianName = "All Phone Numbers"
                    End Select

                    StrCriteriaInput += IIf(StrCriteriaInput = String.Empty, "", ";") + strCriterianName + ";" + strCriterianValue
                Next

                dtReturn = oAdvancedSearchDB.GetResults(strSearchType, StrCriteriaInput, strTankstatus, nLustStatus)
                If dtReturn Is Nothing Then
                    RaiseEvent AdvancedSearchErr("No Matching Results.", Me.ToString)
                End If
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Sub ChangeOrder(ByVal nCurrentOrder As Integer, ByRef dtFavSearchChildList As DataTable, ByVal nDirection As Integer)
            Try
                Dim drow As DataRow
                Dim nOrder As Integer
                Dim xChildInfo As MUSTER.Info.FavSearchChildInfo
                Dim nIndex As Long
                Dim arrKeys(1) As String

                If nDirection = -1 Then
                    If dtFavSearchChildList.Rows.Count <= 0 Then
                        RaiseEvent AdvancedSearchErr("Enter Search Bys / Look Fors Filters.", Me.ToString)
                        Exit Sub
                    End If
                    If dtFavSearchChildList.Rows.Count = 1 Then
                        RaiseEvent AdvancedSearchErr("Insufficient Records to Move Up.", Me.ToString)
                        Exit Sub
                    End If
                    If nCurrentOrder = 1 Then
                        RaiseEvent AdvancedSearchErr("Please select other than First record.", Me.ToString)
                        Exit Sub
                    End If
                Else
                    If dtFavSearchChildList.Rows.Count <= 0 Then
                        RaiseEvent AdvancedSearchErr("Enter Search Bys / Look Fors Filters.", Me.ToString)
                        Exit Sub
                    End If
                    If dtFavSearchChildList.Rows.Count = 1 Then
                        RaiseEvent AdvancedSearchErr("Insufficient Records to Move Down.", Me.ToString)
                        Exit Sub
                    End If

                    If nCurrentOrder = dtFavSearchChildList.Rows.Count Then
                        RaiseEvent AdvancedSearchErr("Please select other than last record.", Me.ToString)
                        Exit Sub
                    End If
                End If
                For Each drow In dtFavSearchChildList.Rows
                    Select Case drow.Item("CRITERION_ORDER")
                        Case nCurrentOrder + nDirection
                            drow.Item("CRITERION_ORDER") = nCurrentOrder
                            xChildInfo = Me.RetrieveChildByName(drow.Item("CRITERION_NAME"))
                            xChildInfo.Order = nCurrentOrder
                        Case nCurrentOrder
                            drow.Item("CRITERION_ORDER") = nCurrentOrder + nDirection
                            xChildInfo = Me.RetrieveChildByName(drow.Item("CRITERION_NAME"))
                            xChildInfo.Order = drow.Item("CRITERION_ORDER")
                        Case Else
                            'Nothing 
                    End Select
                Next
                dtFavSearchChildList.AcceptChanges()
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#End Region
#Region "ExternalEvents"
        Private Sub EnableDelete(ByVal ncount As Integer, ByVal strSrc As String) Handles MyBase.eEnableDelete
            RaiseEvent eEnableDisableDelete(ncount, strSrc)
        End Sub
#End Region
    End Class
End Namespace

