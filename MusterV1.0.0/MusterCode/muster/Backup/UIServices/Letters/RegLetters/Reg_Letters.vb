Imports Microsoft.VisualBasic.FileSystem
Imports System.Data.SqlClient
Imports Utils.DBUtils

Public Class Reg_Letters
    '-------------------------------------------------------------------------------
    ' MUSTER.MUSTER.ShowComment.vb
    '   Provides the mechanism for displaying comments for the app.
    '
    ' Copyright (C) 2004 CIBER, Inc.
    ' All rights reserved.
    '
    ' Release   Initials    Date        Description
    '  1.0        ??      8/??/04    Original class definition.
    '  1.1        JC      1/02/04    Changed AppUser.UserName to AppUser.ID to
    '                                  accomodate new use of pUser by application.
    ' 2.0  Thomas Franey  4/17/09    Added Greeting logic (Mr, Ms, etc) to letter greetings
    ' 1.2.1.3   Hua Cao   07/20/10   Removed strInfoNeeded content for licensee letter
    '                                 
    '-------------------------------------------------------------------------------

    Inherits LetterGenerator
    Private WithEvents cntReconciliation As ContactReconciliation
    Private WithEvents certMail As CertifiedMail
    Private strCertifiedMail As String = String.Empty
    Private nModuleID As Integer = 0
    Private nTitle As String = String.Empty

    Public Sub New()
    End Sub
    Public Sub New(ByVal frmTemp As Form)
        MyBase.New(frmTemp)
    End Sub

#Region "Misc"
    Private Sub CreateColumns(ByVal dtTable As DataTable)
        Try
            dtTable.Columns.Add("OLD CONTACT")
            dtTable.Columns.Add("OLD COMPANY NAME")
            dtTable.Columns.Add("OLD ENTITY")
            dtTable.Columns.Add("OLD MODULE")
            dtTable.Columns.Add("OLD TYPE")
            dtTable.Columns.Add("OLD MOD_BY")
            dtTable.Columns.Add("NEW CONTACT")
            dtTable.Columns.Add("NEW COMPANY NAME")
            dtTable.Columns.Add("NEW ENTITY")
            dtTable.Columns.Add("NEW MODULE")
            dtTable.Columns.Add("NEW TYPE")
            dtTable.Columns.Add("NEW MOD_BY")
            dtTable.Columns.Add("OLD_ADDRESS")
            dtTable.Columns.Add("NEW_ADDRESS")
            dtTable.Columns.Add("ACCEPT", Type.GetType("System.Boolean"))
            dtTable.Columns.Add("REJECT", Type.GetType("System.Boolean"))
            dtTable.Columns.Add("RECONCILIATIONID")
            dtTable.Columns.Add("ENTITYID")
            dtTable.Columns.Add("ENTITYTYPE")
            dtTable.Columns.Add("MODULEID")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub



    Private Function GetXHAndXLContacts(ByVal nEntityID As Integer, ByVal nEntityType As Integer, ByVal moduleID As Integer, Optional ByVal nOwnerID As Integer = 0, Optional ByVal nActive As Integer = -1, Optional ByVal professionalContact As Boolean = False) As DataTable
        Dim dtTable As DataTable
        Dim i As Integer = 0
        Dim pConstruct As New MUSTER.BusinessLogic.pContactStruct
        Dim drRow As DataRow
        Dim cnt As Integer = 0
        Dim dtRecon As New DataTable
        Dim dsReconciliation As DataSet
        Dim drRecon As DataRow
        Try
            dtTable = pConstruct.GETContactName(nEntityID, nEntityType, moduleID, nOwnerID, nActive)
            If dtTable.Rows.Count > 0 Then
                CreateColumns(dtRecon)


                For Each drRow In dtTable.Rows

                    If drRow.Item("Last_Name") > String.Empty And cnt = 0 Then


                        drRow.Item("Greeting") = String.Format("{0} {1}", drRow.Item("First_Name"), drRow.Item("Last_Name"))


                    End If


                    GoTo skip
                    If Not dsReconciliation Is Nothing Then
                        dsReconciliation = Nothing
                    End If
                    dsReconciliation = pConstruct.getReconciliation(Integer.Parse(drRow.Item("ContactAssocID")))
                    If dsReconciliation.Tables.Count > 0 Then

                        Dim dr As DataRow

                        If dsReconciliation.Tables(0).Rows.Count > 0 Then
                            For Each drRecon In dsReconciliation.Tables(0).Rows
                                If drRecon.Item("ENTITYID") = nEntityID And drRecon.Item("ENTITYTYPE") = nEntityType And drRecon.Item("MODULEID") = moduleID Then
                                    dr = dtRecon.NewRow
                                    dr.Item("OLD CONTACT") = drRecon.Item("OLD CONTACT")
                                    dr.Item("OLD COMPANY NAME") = drRecon.Item("OLD COMPANY NAME")
                                    dr.Item("OLD ENTITY") = drRecon.Item("OLD ENTITY")
                                    dr.Item("OLD MODULE") = drRecon.Item("OLD MODULE")
                                    dr.Item("OLD TYPE") = drRecon.Item("OLD TYPE")
                                    dr.Item("OLD MOD_BY") = drRecon.Item("OLD MOD_BY")
                                    dr.Item("NEW CONTACT") = drRecon.Item("NEW CONTACT")
                                    dr.Item("NEW COMPANY NAME") = drRecon.Item("NEW COMPANY NAME")
                                    dr.Item("NEW ENTITY") = drRecon.Item("NEW ENTITY")
                                    dr.Item("NEW MODULE") = drRecon.Item("NEW MODULE")
                                    dr.Item("NEW TYPE") = drRecon.Item("NEW TYPE")
                                    dr.Item("NEW MOD_BY") = drRecon.Item("NEW MOD_BY")
                                    dr.Item("OLD_ADDRESS") = drRecon.Item("OLD_ADDRESS")
                                    dr.Item("NEW_ADDRESS") = drRecon.Item("NEW_ADDRESS")
                                    dr.Item("ACCEPT") = drRecon.Item("ACCEPT")
                                    dr.Item("REJECT") = drRecon.Item("REJECT")
                                    dr.Item("RECONCILIATIONID") = drRecon.Item("RECONCILIATIONID")
                                    dr.Item("ENTITYID") = drRecon.Item("ENTITYID")
                                    dr.Item("ENTITYTYPE") = drRecon.Item("ENTITYTYPE")
                                    dr.Item("MODULEID") = drRecon.Item("MODULEID")
                                    dtRecon.Rows.Add(dr)
                                End If
                            Next
                        End If
                    End If
skip:


                    cnt = cnt + 1
                Next
                If dtRecon.Rows.Count > 0 Then
                    cntReconciliation = New ContactReconciliation(pConstruct, True, dtRecon, nEntityID, nEntityType, moduleID)
                    cntReconciliation.ShowDialog()
                    cntReconciliation.BringToFront()
                    dtTable.Clear()
                    dtTable = pConstruct.GETContactName(nEntityID, nEntityType, moduleID, nOwnerID)


                End If
            End If
            Return dtTable
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Function SetCursorType(ByVal curs As Cursor)
        Try
            If Not IsNothing(frmCaller) Then
                frmCaller.Cursor.Current = curs
            End If
        Catch ex As Exception

        End Try
    End Function
    Private Enum EnumContactType
        XH = 1185
        XL = 1186
    End Enum
    Private Sub FillCCList(ByVal nEntityID As Integer, ByVal moduleID As Integer, ByVal colParams As Specialized.NameValueCollection)
        Dim pContact As New MUSTER.BusinessLogic.pContactStruct
        Try
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'CC List
            Dim dsContacts As DataSet
            Dim dr As DataRow
            Dim strCcList As String = String.Empty
            dsContacts = pContact.GetContactsByEntityAndModule(nEntityID, moduleID)
            If Not dsContacts Is Nothing Then
                If dsContacts.Tables.Count > 0 Then
                    If dsContacts.Tables(0).Rows.Count > 0 Then
                        For Each dr In dsContacts.Tables(0).Rows
                            If Not dr.Item("CC_Info") Is System.DBNull.Value Then
                                If dr.Item("CC_Info") = "YES" Then
                                    If Not dr.Item("DISPLAYAS") Is System.DBNull.Value Then
                                        If dr.Item("DISPLAYAS") <> String.Empty Then
                                            strCcList += dr.Item("DISPLAYAS") + ", "
                                        Else
                                            strCcList += dr.Item("Contact_Name") + ", "
                                        End If
                                    Else
                                        strCcList += dr.Item("Contact_Name") + ", "
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
            End If

            If strCcList <> String.Empty Then
                strCcList = strCcList.Substring(0, strCcList.Length - 2)
                colParams.Add("<CC List>", "CC: " + strCcList)
                colParams.Add("<CList>", strCcList)
            Else
                colParams.Add("<CC List>", String.Empty)
                colParams.Add("<CList>", String.Empty)
            End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Function GetEracIracRep(ByVal nEntityID As Integer, ByVal moduleID As Integer, ByVal strValue As String) As String
        Dim pContact As New MUSTER.BusinessLogic.pContactStruct
        Dim strReturnVal As String = String.Empty
        Dim dsContacts As DataSet
        Try
            dsContacts = pContact.GetContactsByEntityAndModule(nEntityID, moduleID)

            If Not dsContacts Is Nothing Then
                If dsContacts.Tables.Count > 0 Then
                    If dsContacts.Tables(0).Rows.Count > 0 Then
                        For Each dr As DataRow In dsContacts.Tables(0).Rows
                            If Not dr.Item("Type") Is System.DBNull.Value Then
                                If dr.Item("Type").ToString.IndexOf(strValue) >= 0 And dr.Item("IsPerson") Then
                                    If Not dr.Item("DISPLAYAS") Is System.DBNull.Value Then
                                        If dr.Item("DISPLAYAS") <> String.Empty Then
                                            strReturnVal += dr.Item("DISPLAYAS") + ", "
                                        Else
                                            strReturnVal += dr.Item("Contact_Name") + ", "
                                        End If
                                    Else
                                        strReturnVal += dr.Item("Contact_Name") + ", "
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
            End If
            Return strReturnVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "Registration Letters"
    Friend Function GenerateSignatureRequiredLetter(ByVal OwnerID As Integer, ByVal colFacs As MUSTER.Info.FacilityCollection, Optional ByRef own As MUSTER.BusinessLogic.pOwner = Nothing, Optional ByVal bol2ndSigReqLetter As Boolean = False) As Boolean
        Dim strDOC_NAME As String = String.Empty '"REG_SIGN_NF_" + CStr(Trim(nOwnerID.ToString)) + ".doc"
        Dim strToday As String = String.Empty
        Dim colParams As New Specialized.NameValueCollection
        Dim oAddress As New MUSTER.BusinessLogic.pAddress
        Dim dr As DataRow
        Dim dtFacs As New DataTable

        nModuleID = UIUtilsGen.ModuleID.Registration

        Try
            If DOC_PATH = "\" Then
                Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
            End If

            strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm"))
            If bol2ndSigReqLetter Then
                strDOC_NAME = "REG_2nd_SIGN_NF_" + OwnerID.ToString + "_" + strToday + ".doc"
            Else
                strDOC_NAME = "REG_SIGN_NF_" + OwnerID.ToString + "_" + strToday + ".doc"
            End If

            'To Avoid duplicate creation of Letters
            If FileExists(DOC_PATH + strDOC_NAME) Then
                MsgBox(IIf(bol2ndSigReqLetter, "2nd ", "") + "Signature Required Letter - " + strDOC_NAME + "already exists")
                Exit Function
            Else
                'Build NameValueCollection with Tags and Values.
                colParams.Add("<Title>", IIf(bol2ndSigReqLetter, "2nd Signature Needed Letter", "Signature Needed Letter"))
                colParams.Add("<Date>", Format(Now, "MMMM d, yyyy")) ' Format(Now, "MMMM dd, yyyy"))

                'Contact Name
                Dim dtContacts As DataTable = GetXHAndXLContacts(OwnerID, UIUtilsGen.EntityTypes.Owner, UIUtilsGen.ModuleID.Registration, , 1, True)
                If dtContacts.Rows.Count > 0 Then
                    Dim strContactName As String = dtContacts.Rows(0).Item("Greeting")
                    If strContactName <> String.Empty Then
                        colParams.Add("<Contact Name>", strContactName)
                        colParams.Add("<Owner Greeting>", strContactName)
                    Else
                        colParams.Add("<Contact Name>", "")
                    End If
                Else
                    colParams.Add("<Contact Name>", "")
                End If

                If own.OrganizationID > 0 Then
                    colParams.Add("<Owner Name>", own.BPersona.Company)
                    If colParams.Item("<Owner Greeting>") Is Nothing Then
                        colParams.Add("<Owner Greeting>", own.BPersona.Company.Trim)
                    Else
                        If colParams.Item("<Owner Greeting>") = String.Empty Then
                            colParams.Add("<Owner Greeting>", own.BPersona.Company.Trim)
                        End If
                    End If
                Else
                    colParams.Add("<Owner Name>", own.BPersona.Title.Trim & IIf(own.BPersona.Title.Length > 0, " ", "") & own.BPersona.FirstName.Trim & IIf(own.BPersona.MiddleName.Trim.Length > 0, " ", "") & own.BPersona.MiddleName.Trim & " " & own.BPersona.LastName.Trim & IIf(own.BPersona.Suffix.Trim.Length > 0, " ", "") & own.BPersona.Suffix.Trim)
                    If colParams.Item("<Owner Greeting>") Is Nothing Then
                        colParams.Add("<Owner Greeting>", own.BPersona.FirstName.Trim & IIf(own.BPersona.MiddleName.Trim.Length > 0, " ", "") & own.BPersona.MiddleName.Trim & IIf(own.BPersona.LastName.Trim.Length > 0, " ", "") & own.BPersona.LastName.Trim & IIf(own.BPersona.Suffix.Trim.Length > 0, " ", "") & own.BPersona.Suffix.Trim)
                    Else
                        If colParams.Item("<Owner Greeting>") = String.Empty Then
                            colParams.Add("<Owner Greeting>", own.BPersona.FirstName.Trim & IIf(own.BPersona.MiddleName.Trim.Length > 0, " ", "") & own.BPersona.MiddleName.Trim & IIf(own.BPersona.LastName.Trim.Length > 0, " ", "") & own.BPersona.LastName.Trim & IIf(own.BPersona.Suffix.Trim.Length > 0, " ", "") & own.BPersona.Suffix.Trim)
                        End If
                    End If
                End If

                colParams.Add("<Owner Address 1>", own.Address.AddressLine1)
                If own.Address.AddressLine2 = String.Empty Then
                    colParams.Add("<Owner Address 2>", own.Address.City & ", " & own.Address.State.TrimEnd & " " & own.Address.Zip)
                    colParams.Add("<Owner City/State/Zip>", "")
                Else
                    colParams.Add("<Owner Address 2>", own.Address.AddressLine2)
                    colParams.Add("<Owner City/State/Zip>", own.Address.City & ", " & own.Address.State.TrimEnd & " " & own.Address.Zip)
                End If

                dtFacs.Columns.Add("Name")
                dtFacs.Columns.Add("Address")
                dtFacs.Columns.Add("CITY")
                dtFacs.Columns.Add("STATE")
                dtFacs.Columns.Add("ZIP")
                dtFacs.Columns.Add("ID")

                For Each facInfo As MUSTER.Info.FacilityInfo In colFacs.Values
                    If facInfo.ID > 0 Then
                        dr = dtFacs.NewRow
                        dr("Name") = facInfo.Name.Trim
                        oAddress.Retrieve(facInfo.ID)
                        dr("Address") = oAddress.AddressLine1.Trim + IIf(oAddress.AddressLine2.Trim.Length > 0, ", " + oAddress.AddressLine2.Trim, "")
                        dr("CITY") = oAddress.City.Trim
                        dr("STATE") = oAddress.State.Trim
                        dr("ZIP") = oAddress.Zip.Trim
                        dr("ID") = facInfo.ID.ToString
                        dtFacs.Rows.Add(dr)
                    End If
                Next

                If bol2ndSigReqLetter Then
                    strCertifiedMail = String.Empty
                    certMail = New CertifiedMail
                    certMail.ShowDialog()
                    If strCertifiedMail = String.Empty Then
                        colParams.Add("<Certified Mail>", "")
                    Else
                        colParams.Add("<Certified Mail>", strCertifiedMail)
                    End If
                End If

                colParams.Add("<Date Required>", DateAdd(DateInterval.Day, 30, Today.Date).ToShortDateString)
                colParams.Add("<User Phone>", CType(MusterContainer.AppUser.PhoneNumber, String))
                colParams.Add("<User>", MusterContainer.AppUser.Name)
                FillCCList(OwnerID, UIUtilsGen.ModuleID.Registration, colParams)

                Dim oWord As Word.Application = MusterContainer.GetWordApp

                If Not oWord Is Nothing Then
                    Dim strTempPath As String = TmpltPath + "Registration\SignatureNeededLetter.doc"
                    Dim ltrGen As New MUSTER.BusinessLogic.pLetterGen
                    ltrGen.CreateOtherRegistrationLetter(OwnerID, colParams, strTempPath, Doc_Path + strDOC_NAME, dtFacs, own, oWord)
                    UIUtilsGen.SaveDocument(OwnerID, UIUtilsGen.EntityTypes.Owner, strDOC_NAME, "Sig Needed Letter", DOC_PATH, "Signature is required for one or more facilities", nModuleID, 0, 0, 0)
                    oWord.Visible = True
                End If

                oWord = Nothing


                End If

        Catch ex As Exception
            Throw ex
        Finally
            SetCursorType(System.Windows.Forms.Cursors.Default)
        End Try
    End Function
    Friend Function GenerateUpcomingInstallLetter(ByVal OwnerID As Integer, ByVal colFacs As MUSTER.Info.FacilityCollection, Optional ByRef own As MUSTER.BusinessLogic.pOwner = Nothing) As Boolean
        Dim strDOC_NAME As String = String.Empty
        Dim strToday As String = String.Empty
        Dim colParams As New Specialized.NameValueCollection
        Dim oAddress As New MUSTER.BusinessLogic.pAddress
        Dim dr As DataRow
        Dim dtFacs As New DataTable
        Dim pLicensee As New MUSTER.BusinessLogic.pLicensee
        nModuleID = UIUtilsGen.ModuleID.Registration

        Try
            If DOC_PATH = "\" Then
                Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
            End If

            strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm"))
            strDOC_NAME = "REG_Upcoming_Install_" + OwnerID.ToString + "_" + strToday + ".doc"

            'To Avoid duplicate creation of Letters
            If FileExists(DOC_PATH + strDOC_NAME) Then
                MsgBox("Signature Required Letter - " + strDOC_NAME + "already exists")
                Exit Function
            Else
                'Build NameValueCollection with Tags and Values.
                colParams.Add("<Title>", "Upcoming Install Letter")
                colParams.Add("<Date>", Format(Now, "MMMM d, yyyy"))

                'Contact Name
                Dim dtContacts As DataTable = GetXHAndXLContacts(OwnerID, UIUtilsGen.EntityTypes.Owner, UIUtilsGen.ModuleID.Registration, , 1)
                If dtContacts.Rows.Count > 0 Then
                    Dim strContactName As String = dtContacts.Rows(0).Item("Greeting")
                    If strContactName <> String.Empty Then
                        colParams.Add("<Contact Name>", strContactName)
                        colParams.Add("<Owner Greeting>", strContactName)
                    Else
                        colParams.Add("<Contact Name>", "")
                    End If
                Else
                    colParams.Add("<Contact Name>", "")
                End If

                If own.OrganizationID > 0 Then
                    colParams.Add("<Owner Name>", own.BPersona.Company)
                    If colParams.Item("<Owner Greeting>") Is Nothing Then
                        colParams.Add("<Owner Greeting>", own.BPersona.Company.Trim)
                    Else
                        If colParams.Item("<Owner Greeting>") = String.Empty Then
                            colParams.Add("<Owner Greeting>", own.BPersona.Company.Trim)
                        End If
                    End If
                Else
                    colParams.Add("<Owner Name>", own.BPersona.Title.Trim & IIf(own.BPersona.Title.Length > 0, " ", "") & own.BPersona.FirstName.Trim & IIf(own.BPersona.MiddleName.Trim.Length > 0, " ", "") & own.BPersona.MiddleName.Trim & " " & own.BPersona.LastName.Trim & IIf(own.BPersona.Suffix.Trim.Length > 0, " ", "") & own.BPersona.Suffix.Trim)
                    If colParams.Item("<Owner Greeting>") Is Nothing Then
                        colParams.Add("<Owner Greeting>", own.BPersona.FirstName.Trim & IIf(own.BPersona.MiddleName.Trim.Length > 0, " ", "") & own.BPersona.MiddleName.Trim & IIf(own.BPersona.LastName.Trim.Length > 0, " ", "") & own.BPersona.LastName.Trim & IIf(own.BPersona.Suffix.Trim.Length > 0, " ", "") & own.BPersona.Suffix.Trim)
                    Else
                        If colParams.Item("<Owner Greeting>") = String.Empty Then
                            colParams.Add("<Owner Greeting>", own.BPersona.FirstName.Trim & IIf(own.BPersona.MiddleName.Trim.Length > 0, " ", "") & own.BPersona.MiddleName.Trim & IIf(own.BPersona.LastName.Trim.Length > 0, " ", "") & own.BPersona.LastName.Trim & IIf(own.BPersona.Suffix.Trim.Length > 0, " ", "") & own.BPersona.Suffix.Trim)
                        End If
                    End If
                End If

                colParams.Add("<Owner Address 1>", own.Address.AddressLine1)
                If own.Address.AddressLine2 = String.Empty Then
                    colParams.Add("<Owner Address 2>", own.Address.City & ", " & own.Address.State.TrimEnd & " " & own.Address.Zip)
                    colParams.Add("<Owner City/State/Zip>", "")
                Else
                    colParams.Add("<Owner Address 2>", own.Address.AddressLine2)
                    colParams.Add("<Owner City/State/Zip>", own.Address.City & ", " & own.Address.State.TrimEnd & " " & own.Address.Zip)
                End If

                dtFacs.Columns.Add("Name")
                dtFacs.Columns.Add("Address")
                dtFacs.Columns.Add("CITY")
                dtFacs.Columns.Add("STATE")
                dtFacs.Columns.Add("ZIP")
                dtFacs.Columns.Add("ID")

                For Each facInfo As MUSTER.Info.FacilityInfo In colFacs.Values
                    If facInfo.ID > 0 Then
                        If facInfo.LicenseeID > 0 Then
                            pLicensee.Retrieve(facInfo.LicenseeID)
                            colParams.Add("<CList>", pLicensee.Licensee_name.Trim)
                        End If
                        dr = dtFacs.NewRow
                        dr("Name") = facInfo.Name.Trim
                        oAddress.Retrieve(facInfo.AddressID)
                        dr("Address") = oAddress.AddressLine1.Trim + IIf(oAddress.AddressLine2.Trim.Length > 0, ", " + oAddress.AddressLine2.Trim, "")
                        dr("CITY") = oAddress.City.Trim
                        dr("STATE") = oAddress.State.Trim
                        dr("ZIP") = oAddress.Zip.Trim
                        dr("ID") = facInfo.ID.ToString
                        dtFacs.Rows.Add(dr)
                    End If
                Next

                colParams.Add("<Date Required>", DateAdd(DateInterval.Day, 30, Today.Date).ToShortDateString)
                colParams.Add("<User Phone>", CType(MusterContainer.AppUser.PhoneNumber, String))
                colParams.Add("<User>", MusterContainer.AppUser.Name)
                FillCCList(OwnerID, UIUtilsGen.ModuleID.Registration, colParams)

                Dim oWord As Word.Application = MusterContainer.GetWordApp

                If Not oWord Is Nothing Then


                    Dim strTempPath As String = TmpltPath + "Registration\UpcomingInstallLetter.doc"
                    Dim ltrGen As New MUSTER.BusinessLogic.pLetterGen
                    ltrGen.CreateOtherRegistrationLetter(OwnerID, colParams, strTempPath, Doc_Path + strDOC_NAME, dtFacs, own, oWord)
                    UIUtilsGen.SaveDocument(OwnerID, UIUtilsGen.EntityTypes.Owner, strDOC_NAME, "Upcoming Install Letter", DOC_PATH, "Upcoming Install for one or more facilities", nModuleID, 0, 0, 0)
                    oWord.Visible = True
                End If

                oWord = Nothing

            End If


        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Friend Function GenerateRegistrationLetter(ByVal OwnerID As Integer, ByVal slRegActivity As SortedList, ByVal dtFacs As DataTable, ByVal own As MUSTER.BusinessLogic.pOwner, ByVal strTOSITanks As String, ByVal strTransferFacs As String, Optional ByVal showTransferOwnerSection As Boolean = False) As Boolean
        Dim strDOC_NAME As String = String.Empty
        Dim strToday As String = String.Empty
        Dim colParams As New Specialized.NameValueCollection
        nModuleID = UIUtilsGen.ModuleID.Registration
        Try
            If DOC_PATH = "\" Then
                Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
            End If

            strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm"))
            strDOC_NAME = "REG_" + OwnerID.ToString + "_" + strToday + ".doc"

            'To Avoid duplicate creation of Letters
            If FileExists(DOC_PATH + strDOC_NAME) Then
                MsgBox("Registration Letter - " + strDOC_NAME + "already exists")
                Exit Function
            Else
                'Build NameValueCollection with Tags and Values.
                colParams.Add("<Title>", "Registration Letter")
                colParams.Add("<Date>", Format(Now, "MMMM d, yyyy"))

                'Contact Name
                Dim dtContacts As DataTable = GetXHAndXLContacts(OwnerID, UIUtilsGen.EntityTypes.Owner, UIUtilsGen.ModuleID.Registration, , 1)
                If dtContacts.Rows.Count > 0 Then
                    Dim strContactName As String = IIf(dtContacts.Rows(0).Item("Greeting") Is DBNull.Value, "", dtContacts.Rows(0).Item("Greeting"))
                    If strContactName <> String.Empty Then
                        colParams.Add("<Contact Name>", strContactName)
                        colParams.Add("<Owner Greeting>", strContactName)
                    Else
                        colParams.Add("<Contact Name>", "")
                    End If
                Else
                    colParams.Add("<Contact Name>", "")
                End If

                If own.OrganizationID > 0 Then
                    colParams.Add("<Owner Name>", own.BPersona.Company)
                    If colParams.Item("<Owner Greeting>") Is Nothing Then
                        colParams.Add("<Owner Greeting>", own.BPersona.Company.Trim)
                    Else
                        If colParams.Item("<Owner Greeting>") = String.Empty Then
                            colParams.Add("<Owner Greeting>", own.BPersona.Company.Trim)
                        End If
                    End If
                Else
                    colParams.Add("<Owner Name>", own.BPersona.Title.Trim & IIf(own.BPersona.Title.Length > 0, " ", "") & own.BPersona.FirstName.Trim & IIf(own.BPersona.MiddleName.Trim.Length > 0, " ", "") & own.BPersona.MiddleName.Trim & " " & own.BPersona.LastName.Trim & IIf(own.BPersona.Suffix.Trim.Length > 0, " ", "") & own.BPersona.Suffix.Trim)
                    If colParams.Item("<Owner Greeting>") Is Nothing Then
                        colParams.Add("<Owner Greeting>", own.BPersona.FirstName.Trim & IIf(own.BPersona.MiddleName.Trim.Length > 0, " ", "") & own.BPersona.MiddleName.Trim & " " & own.BPersona.LastName.Trim & IIf(own.BPersona.Suffix.Trim.Length > 0, " ", "") & own.BPersona.Suffix.Trim)
                    Else
                        If colParams.Item("<Owner Greeting>") = String.Empty Then
                            colParams.Add("<Owner Greeting>", own.BPersona.FirstName.Trim & IIf(own.BPersona.MiddleName.Trim.Length > 0, " ", "") & own.BPersona.MiddleName.Trim & " " & own.BPersona.LastName.Trim & IIf(own.BPersona.Suffix.Trim.Length > 0, " ", "") & own.BPersona.Suffix.Trim)
                        End If
                    End If
                End If

                colParams.Add("<Owner Address 1>", own.Address.AddressLine1)
                If own.Address.AddressLine2 = String.Empty Then
                    colParams.Add("<Owner Address 2>", own.Address.City & ", " & own.Address.State.TrimEnd & " " & own.Address.Zip)
                    colParams.Add("<Owner City/State/Zip>", "")
                Else
                    colParams.Add("<Owner Address 2>", own.Address.AddressLine2)
                    colParams.Add("<Owner City/State/Zip>", own.Address.City & ", " & own.Address.State.TrimEnd & " " & own.Address.Zip)
                End If

                If slRegActivity.Contains(UIUtilsGen.ActivityTypes.TankStatusTOSI) Then
                    If Not slRegActivity.Item(UIUtilsGen.ActivityTypes.TankStatusTOSI) Is Nothing Then
                        colParams.Add("<TOSI Facility IDs>", strTOSITanks)
                        'colParams.Add("<TOSI Facility IDs>", slRegActivity.Item(UIUtilsGen.ActivityTypes.TankStatusTOSI).ToString)
                    End If
                End If
                If slRegActivity.Contains(UIUtilsGen.ActivityTypes.TransferOwnership) Then
                    If Not slRegActivity.Item(UIUtilsGen.ActivityTypes.TransferOwnership) Is Nothing Then
                        colParams.Add("<Transfer Facility IDs>", strTransferFacs)
                        If strTransferFacs.Split(",").Length > 1 Then
                            colParams.Add("<This facility has / These facilities have>", "These facilities have")
                        Else
                            colParams.Add("<This facility has / These facilities have>", "This facility has")
                        End If
                    End If
                End If

                colParams.Add("<Date Required>", DateAdd(DateInterval.Day, 30, Today.Date).ToShortDateString)
                colParams.Add("<User Phone>", CType(MusterContainer.AppUser.PhoneNumber, String))
                colParams.Add("<User>", MusterContainer.AppUser.Name)
                FillCCList(OwnerID, UIUtilsGen.ModuleID.Registration, colParams)

                Dim oWord As Word.Application = MusterContainer.GetWordApp

                If Not oWord Is Nothing Then


                    Dim strTempPath As String = TmpltPath + "Registration\RegistrationLetter.doc"
                    Dim ltrGen As New MUSTER.BusinessLogic.pLetterGen
                    ltrGen.CreateRegistrationLetter(OwnerID, colParams, strTempPath, Doc_Path + strDOC_NAME, dtFacs, own, oWord, showTransferOwnerSection)
                    UIUtilsGen.SaveDocument(OwnerID, UIUtilsGen.EntityTypes.Owner, strDOC_NAME, "Registration Letter", DOC_PATH, "Registration for one or more facilities", nModuleID, 0, 0, 0)
                    oWord.Visible = True
                End If

                oWord = Nothing
            End If


        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Friend Function GenerateTransferAcknowledgementLetter(ByVal OwnerID As Integer, ByVal colFacs As MUSTER.Info.FacilityCollection, Optional ByRef own As MUSTER.BusinessLogic.pOwner = Nothing, Optional ByVal isLetterForNewOwner As Boolean = False) As Boolean
        Dim strDOC_NAME As String = String.Empty
        Dim strToday As String = String.Empty
        Dim colParams As New Specialized.NameValueCollection
        Dim oAddress As New MUSTER.BusinessLogic.pAddress
        Dim dr As DataRow
        Dim dtFacs As New DataTable
        nModuleID = UIUtilsGen.ModuleID.Registration
        Try
            If DOC_PATH = "\" Then
                Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
            End If

            strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm"))
            strDOC_NAME = "TransferAcknowledgment_" + OwnerID.ToString + "_" + strToday + ".doc"

            'To Avoid duplicate creation of Letters
            If FileExists(DOC_PATH + strDOC_NAME) Then
                MsgBox("Transfer Acknowledgement Letter - " + strDOC_NAME + "already exists")
                Exit Function
            Else
                'Build NameValueCollection with Tags and Values.
                colParams.Add("<Title>", "Transfer Acknowledgement Letter")
                colParams.Add("<Date>", Format(Now, "MMMM d, yyyy"))

                If own Is Nothing Then
                    own = New MUSTER.BusinessLogic.pOwner
                End If
                If own.ID <> OwnerID Then
                    own.Retrieve(OwnerID, , , True)
                End If

                'Contact Name
                Dim dtContacts As DataTable = GetXHAndXLContacts(OwnerID, UIUtilsGen.EntityTypes.Owner, UIUtilsGen.ModuleID.Registration, , 1)
                If dtContacts.Rows.Count > 0 Then
                    Dim strContactName As String = IIf(dtContacts.Rows(0).Item("Greeting") Is DBNull.Value, "", dtContacts.Rows(0).Item("Greeting"))
                    If strContactName <> String.Empty Then
                        colParams.Add("<Contact Name>", strContactName)
                        colParams.Add("<Owner Greeting>", strContactName)
                    Else
                        colParams.Add("<Contact Name>", "")
                    End If
                Else
                    colParams.Add("<Contact Name>", "")
                End If

                If own.OrganizationID > 0 Then
                    colParams.Add("<Owner Name>", own.BPersona.Company)
                    If colParams.Item("<Owner Greeting>") Is Nothing Then
                        colParams.Add("<Owner Greeting>", own.BPersona.Company.Trim)
                    Else
                        If colParams.Item("<Owner Greeting>") = String.Empty Then
                            colParams.Add("<Owner Greeting>", own.BPersona.Company.Trim)
                        End If
                    End If
                Else
                    colParams.Add("<Owner Name>", own.BPersona.Title.Trim & IIf(own.BPersona.Title.Length > 0, " ", "") & own.BPersona.FirstName.Trim & IIf(own.BPersona.MiddleName.Trim.Length > 0, " ", "") & own.BPersona.MiddleName.Trim & " " & own.BPersona.LastName.Trim & IIf(own.BPersona.Suffix.Trim.Length > 0, " ", "") & own.BPersona.Suffix.Trim)
                    If colParams.Item("<Owner Greeting>") Is Nothing Then
                        colParams.Add("<Owner Greeting>", own.BPersona.FirstName.Trim & IIf(own.BPersona.MiddleName.Trim.Length > 0, " ", "") & own.BPersona.MiddleName.Trim & " " & own.BPersona.LastName.Trim & IIf(own.BPersona.Suffix.Trim.Length > 0, " ", "") & own.BPersona.Suffix.Trim)

                    Else
                        If colParams.Item("<Owner Greeting>") = String.Empty Then
                            colParams.Add("<Owner Greeting>", own.BPersona.FirstName.Trim & IIf(own.BPersona.MiddleName.Trim.Length > 0, " ", "") & own.BPersona.MiddleName.Trim & " " & own.BPersona.LastName.Trim & IIf(own.BPersona.Suffix.Trim.Length > 0, " ", "") & own.BPersona.Suffix.Trim)
                        End If
                    End If
                End If

                colParams.Add("<Owner Address 1>", own.Address.AddressLine1)
                If own.Address.AddressLine2 = String.Empty Then
                    colParams.Add("<Owner Address 2>", own.Address.City & ", " & own.Address.State.TrimEnd & " " & own.Address.Zip)
                    colParams.Add("<Owner City/State/Zip>", "")
                Else
                    colParams.Add("<Owner Address 2>", own.Address.AddressLine2)
                    colParams.Add("<Owner City/State/Zip>", own.Address.City & ", " & own.Address.State.TrimEnd & " " & own.Address.Zip)
                End If

                dtFacs.Columns.Add("Name")
                dtFacs.Columns.Add("Address")
                dtFacs.Columns.Add("CITY")
                dtFacs.Columns.Add("STATE")
                dtFacs.Columns.Add("ZIP")
                dtFacs.Columns.Add("ID")

                For Each facInfo As MUSTER.Info.FacilityInfo In colFacs.Values
                    If facInfo.ID > 0 Then
                        dr = dtFacs.NewRow
                        dr("Name") = facInfo.Name.Trim
                        oAddress.Retrieve(facInfo.AddressID)
                        dr("Address") = oAddress.AddressLine1.Trim + IIf(oAddress.AddressLine2.Trim.Length > 0, ", " + oAddress.AddressLine2.Trim, "")
                        dr("CITY") = oAddress.City.Trim
                        dr("STATE") = oAddress.State.Trim
                        dr("ZIP") = oAddress.Zip.Trim
                        dr("ID") = facInfo.ID.ToString
                        dtFacs.Rows.Add(dr)
                    End If
                Next

                colParams.Add("<facility / facilities>", IIf(dtFacs.Rows.Count > 1, "facilities", "facility"))
                colParams.Add("<User Phone>", CType(MusterContainer.AppUser.PhoneNumber, String))
                colParams.Add("<User>", MusterContainer.AppUser.Name)
                FillCCList(OwnerID, UIUtilsGen.ModuleID.Registration, colParams)

                Dim oWord As Word.Application = MusterContainer.GetWordApp

                If Not oWord Is Nothing Then


                    Dim strTempPath As String
                    If isLetterForNewOwner Then
                        strTempPath = TmpltPath + "Registration\TransferAcknowledgementLetter_NewOwner.doc"
                    Else
                        strTempPath = TmpltPath + "Registration\TransferAcknowledgementLetter.doc"
                    End If
                    Dim ltrGen As New MUSTER.BusinessLogic.pLetterGen
                    ltrGen.CreateOtherRegistrationLetter(OwnerID, colParams, strTempPath, Doc_Path + strDOC_NAME, dtFacs, own, oWord)
                    UIUtilsGen.SaveDocument(OwnerID, UIUtilsGen.EntityTypes.Owner, strDOC_NAME, "Transfer Acknowledgement Letter", DOC_PATH, "Transfer Acknowledgement Letter for transferring one or more facilities", nModuleID, 0, 0, 0)
                    oWord.Visible = True
                End If
                oWord = Nothing

            End If


        Catch ex As Exception
            Throw ex
        Finally
            SetCursorType(System.Windows.Forms.Cursors.Default)
        End Try
    End Function
    Friend Sub GenerateComplianceLetter(ByVal isNonComplianceLetter As Boolean, ByVal nFacID As String, ByVal strFiscalYear As String)
        Dim strDOC_NAME As String = String.Empty '"REG_SIGN_NF_" + CStr(Trim(nOwnerID.ToString)) + ".doc"
        Dim strToday As String = String.Empty
        Dim colParams As New Specialized.NameValueCollection
        Dim oOwn As New MUSTER.BusinessLogic.pOwner
        Dim oFac As New MUSTER.BusinessLogic.pFacility
        Dim oAddInfo As New MUSTER.Info.AddressInfo
        Dim ownerID As Integer = 0
        Dim ltrGen As New MUSTER.BusinessLogic.pLetterGen
        Dim bolFacHasAllTOSITanks As Boolean = False
        Dim facs() As String = Nothing

        nModuleID = UIUtilsGen.ModuleID.Registration
        Try
            If Not nFacID Is Nothing AndAlso nFacID.Length > 0 Then
                facs = nFacID.Split(",")
            End If

            If facs Is Nothing OrElse Not IsNumeric(facs(0)) Then
                Throw New Exception("Improper Facility ID format used.")
            End If

            If DOC_PATH = "\" Then
                Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
            End If

            strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm"))

            If isNonComplianceLetter Then
                strDOC_NAME = "REG_NonComplianceLetter_" + nFacID.ToString + "_" + strToday + ".doc"
            Else
                strDOC_NAME = "REG_ComplianceLetter_" + nFacID.ToString + "_" + strToday + ".doc"
            End If
            'To Avoid duplicate creation of Letters
            If FileExists(DOC_PATH + strDOC_NAME) Then
                MsgBox(IIf(isNonComplianceLetter, "Non ", "") + "Compliance Letter - " + strDOC_NAME + "already exists")
                Exit Sub
            Else
                'Build NameValueCollection with Tags and Values.
                colParams.Add("<Title>", IIf(isNonComplianceLetter, "Non Compliance Letter", "Compliance Letter"))
                colParams.Add("<Date>", Format(Now, "MMMM d, yyyy"))

                oFac.Retrieve(New MUSTER.Info.OwnerInfo, Convert.ToInt32(facs(0)), , "FACILITY", False, True)
                oAddInfo = oFac.FacilityAddress
                oOwn.Retrieve(oFac.OwnerID, , False, True)
                ownerID = oFac.OwnerID

                If oOwn.OrganizationID > 0 Then
                    colParams.Add("<Owner Name>", oOwn.BPersona.Company.Trim)
                    colParams.Add("<Owner Greeting>", "Dear Owner")

                Else
                    colParams.Add("<Owner Name>", IIf(oOwn.BPersona.Title.Trim.Length > 0, oOwn.BPersona.Title.Trim + " ", "") + _
                                                    oOwn.BPersona.FirstName.Trim + _
                                                    IIf(oOwn.BPersona.MiddleName.Trim.Length > 0, oOwn.BPersona.MiddleName.Trim + " ", " ") + _
                                                    oOwn.BPersona.LastName.Trim + IIf(oOwn.BPersona.Suffix.Trim.Length > 0, " ", "") + oOwn.BPersona.Suffix.Trim)

                    colParams.Add("<Owner Greeting>", "Dear " + IIf(oOwn.BPersona.Title.Trim.Length > 0, oOwn.BPersona.Title.Trim + " ", "") + _
                                                  oOwn.BPersona.FirstName.Trim + _
                                                  IIf(oOwn.BPersona.MiddleName.Trim.Length > 0, oOwn.BPersona.MiddleName.Trim + " ", " ") + _
                                                  oOwn.BPersona.LastName.Trim + IIf(oOwn.BPersona.Suffix.Trim.Length > 0, " ", "") + oOwn.BPersona.Suffix.Trim)
                End If

                colParams.Add("<Owner Address 1>", oOwn.Address.AddressLine1)
                If oOwn.Address.AddressLine2.Trim = String.Empty Then
                    colParams.Add("<Owner Address 2>", oOwn.Address.City & ", " & oOwn.Address.State.TrimEnd & " " & oOwn.Address.Zip)
                    colParams.Add("<Owner City/State/Zip>", "<DeleteMe>")
                Else
                    colParams.Add("<Owner Address 2>", oOwn.Address.AddressLine2)
                    colParams.Add("<Owner City/State/Zip>", oOwn.Address.City & ", " & oOwn.Address.State.TrimEnd & " " & oOwn.Address.Zip)
                End If

                Dim facstring As String = String.Empty

                For Each fac As String In facs
                    If IsNumeric(fac) Then
                        oFac.Retrieve(New MUSTER.Info.OwnerInfo, Convert.ToInt32(fac), , "FACILITY", False, True)
                        oAddInfo = oFac.FacilityAddress

                        If oAddInfo.AddressLine2.Trim.Length > 0 Then
                            facstring = String.Format("{0}{1}ID # {2} - {3}, {4} {5}   {6},{7}", facstring, IIf(facstring.Length > 0, vbCrLf, String.Empty), _
                                                      fac, oFac.Name, oAddInfo.AddressLine1.Trim, oAddInfo.AddressLine2.Trim, oAddInfo.City, oAddInfo.State)
                        Else
                            facstring = String.Format("{0}{1} ID # {2} - {3}, {4}  {5},{6}", facstring, IIf(facstring.Length > 0, vbCrLf, String.Empty), _
                                                      fac, oFac.Name, oAddInfo.AddressLine1.Trim, oAddInfo.City, oAddInfo.State)
                        End If

                        Dim ds As DataSet = oOwn.RunSQLQuery("SELECT (SELECT COUNT(*) FROM TBLREG_TANK WHERE FACILITY_ID = " + fac + " AND DELETED = 0) AS TOTAL, (SELECT COUNT(*) FROM TBLREG_TANK WHERE FACILITY_ID = " + fac + " AND DELETED = 0 AND TANKSTATUS = 429) AS TOSI")
                        If ds.Tables.Count > 0 Then
                            If ds.Tables(0).Rows.Count > 0 Then
                                If ds.Tables(0).Rows(0)(0) = ds.Tables(0).Rows(0)(1) Then
                                    bolFacHasAllTOSITanks = True
                                End If
                            End If
                        End If
                    End If
                Next

                colParams.Add("<facility>", facstring)

                colParams.Add("<Fiscal Year>", strFiscalYear)

                If facs.GetUpperBound(0) = 0 Then
                    colParams.Add("<this facility is/these facilities are>", "this facility is")
                    colParams.Add("<facility/facilities>", "facility")
                Else
                    colParams.Add("<this facility is/these facilities are>", "these facilities are")
                    colParams.Add("<facility/facilities>", "facilities")
                End If


                Dim oWord As Word.Application = MusterContainer.GetWordApp

                If Not oWord Is Nothing Then




                    Dim strTempPath As String
                    If isNonComplianceLetter Then
                        strTempPath = TmpltPath + "Registration\NonComplianceLetter.doc"
                    Else
                        strTempPath = TmpltPath + "Registration\ComplianceLetter.doc"
                    End If
                    ltrGen.CreateComplianceLetter(colParams, strTempPath, doc_path + strDOC_NAME, oWord, bolFacHasAllTOSITanks)
                    UIUtilsGen.SaveDocument(ownerID, UIUtilsGen.EntityTypes.Owner, strDOC_NAME, _
                                            IIf(isNonComplianceLetter, "Non Compliance Letter", "Compliance Letter"), DOC_PATH, _
                                            IIf(isNonComplianceLetter, "Non Compliance Letter", "Compliance Letter"), UIUtilsGen.ModuleID.Registration, 0, 0, 0)
                    oWord.Visible = True
                End If
                oWord = Nothing
            End If

        Catch ex As Exception
            Throw ex
        Finally
            SetCursorType(System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Friend Sub GenerateTOSILetter(ByVal isNonComplianceLetter As Boolean, ByVal nFacID As String)
        Dim strDOC_NAME As String = String.Empty '"REG_SIGN_NF_" + CStr(Trim(nOwnerID.ToString)) + ".doc"
        Dim strToday As String = String.Empty
        Dim colParams As New Specialized.NameValueCollection
        Dim oOwn As New MUSTER.BusinessLogic.pOwner
        Dim oFac As New MUSTER.BusinessLogic.pFacility
        Dim oAddInfo As New MUSTER.Info.AddressInfo
        Dim ownerID As Integer = 0
        Dim ltrGen As New MUSTER.BusinessLogic.pLetterGen
        Dim bolFacHasAllTOSITanks As Boolean = False
        Dim facs() As String = Nothing

        nModuleID = UIUtilsGen.ModuleID.Registration
        Try
            If Not nFacID Is Nothing AndAlso nFacID.Length > 0 Then
                facs = nFacID.Split(",")
            End If

            If facs Is Nothing OrElse Not IsNumeric(facs(0)) Then
                Throw New Exception("Improper Facility ID format used.")
            End If

            If DOC_PATH = "\" Then
                Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
            End If

            strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm"))

           
            strDOC_NAME = "REG_TOSILetter_" + nFacID.ToString + "_" + strToday + ".doc"

            'To Avoid duplicate creation of Letters
            If FileExists(DOC_PATH + strDOC_NAME) Then
                MsgBox(IIf(isNonComplianceLetter, "Non ", "") + "Compliance Letter - " + strDOC_NAME + "already exists")
                Exit Sub
            Else
                'Build NameValueCollection with Tags and Values.
                colParams.Add("<Title>", IIf(isNonComplianceLetter, "Non Compliance Letter", "Compliance Letter"))
                colParams.Add("<Date>", Format(Now, "MMMM d, yyyy"))

                oFac.Retrieve(New MUSTER.Info.OwnerInfo, Convert.ToInt32(facs(0)), , "FACILITY", False, True)
                oAddInfo = oFac.FacilityAddress
                oOwn.Retrieve(oFac.OwnerID, , False, True)
                ownerID = oFac.OwnerID

                If oOwn.OrganizationID > 0 Then
                    colParams.Add("<Owner Name>", oOwn.BPersona.Company.Trim)
                    colParams.Add("<Owner Greeting>", "Dear Owner")

                Else
                    colParams.Add("<Owner Name>", IIf(oOwn.BPersona.Title.Trim.Length > 0, oOwn.BPersona.Title.Trim + " ", "") + _
                                                    oOwn.BPersona.FirstName.Trim + _
                                                    IIf(oOwn.BPersona.MiddleName.Trim.Length > 0, oOwn.BPersona.MiddleName.Trim + " ", " ") + _
                                                    oOwn.BPersona.LastName.Trim + IIf(oOwn.BPersona.Suffix.Trim.Length > 0, " ", "") + oOwn.BPersona.Suffix.Trim)

                    colParams.Add("<Owner Greeting>", "Dear " + IIf(oOwn.BPersona.Title.Trim.Length > 0, oOwn.BPersona.Title.Trim + " ", "") + _
                                                  oOwn.BPersona.FirstName.Trim + _
                                                  IIf(oOwn.BPersona.MiddleName.Trim.Length > 0, oOwn.BPersona.MiddleName.Trim + " ", " ") + _
                                                  oOwn.BPersona.LastName.Trim + IIf(oOwn.BPersona.Suffix.Trim.Length > 0, " ", "") + oOwn.BPersona.Suffix.Trim)
                End If

                colParams.Add("<Owner Address 1>", oOwn.Address.AddressLine1)
                If oOwn.Address.AddressLine2.Trim = String.Empty Then
                    colParams.Add("<Owner Address 2>", oOwn.Address.City & ", " & oOwn.Address.State.TrimEnd & " " & oOwn.Address.Zip)
                    colParams.Add("<Owner City/State/Zip>", "<DeleteMe>")
                Else
                    colParams.Add("<Owner Address 2>", oOwn.Address.AddressLine2)
                    colParams.Add("<Owner City/State/Zip>", oOwn.Address.City & ", " & oOwn.Address.State.TrimEnd & " " & oOwn.Address.Zip)
                End If

                Dim facstring As String = String.Empty

                For Each fac As String In facs
                    If IsNumeric(fac) Then
                        oFac.Retrieve(New MUSTER.Info.OwnerInfo, Convert.ToInt32(fac), , "FACILITY", False, True)
                        oAddInfo = oFac.FacilityAddress

                        If oAddInfo.AddressLine2.Trim.Length > 0 Then
                            facstring = String.Format("{0}{1}ID # {2} - {3}, {4} {5}   {6}, {7}", facstring, IIf(facstring.Length > 0, vbCrLf, String.Empty), _
                                                      fac, oFac.Name, oAddInfo.AddressLine1.Trim, oAddInfo.AddressLine2.Trim, oAddInfo.City, oAddInfo.State)
                        Else
                            facstring = String.Format("{0}{1} ID # {2} - {3}, {4}  {5}, {6}", facstring, IIf(facstring.Length > 0, vbCrLf, String.Empty), _
                                                      fac, oFac.Name, oAddInfo.AddressLine1.Trim, oAddInfo.City, oAddInfo.State)
                        End If

                        Dim ds As DataSet = oOwn.RunSQLQuery("SELECT (SELECT COUNT(*) FROM TBLREG_TANK WHERE FACILITY_ID = " + fac + " AND DELETED = 0) AS TOTAL, (SELECT COUNT(*) FROM TBLREG_TANK WHERE FACILITY_ID = " + fac + " AND DELETED = 0 AND TANKSTATUS = 429) AS TOSI")
                        If ds.Tables.Count > 0 Then
                            If ds.Tables(0).Rows.Count > 0 Then
                                If ds.Tables(0).Rows(0)(0) = ds.Tables(0).Rows(0)(1) Then
                                    bolFacHasAllTOSITanks = True
                                End If
                            End If
                        End If
                    End If
                Next

                colParams.Add("<facility>", facstring)

                '  colParams.Add("<Fiscal Year>", strFiscalYear)

                If facs.GetUpperBound(0) = 0 Then
                    colParams.Add("<this facility is/these facilities are>", "this facility is")
                    colParams.Add("<facility/facilities>", "facility")
                Else
                    colParams.Add("<this facility is/these facilities are>", "these facilities are")
                    colParams.Add("<facility/facilities>", "facilities")
                End If
                colParams.Add("<User Phone>", CType(MusterContainer.AppUser.PhoneNumber, String))
                colParams.Add("<User>", MusterContainer.AppUser.Name)
                Dim oWord As Word.Application = MusterContainer.GetWordApp

                If Not oWord Is Nothing Then




                    Dim strTempPath As String

                    strTempPath = TmpltPath + "Registration\TOS-I.doc"

                    ltrGen.CreateTOSILetter(colParams, strTempPath, doc_path + strDOC_NAME, oWord, bolFacHasAllTOSITanks)
                    UIUtilsGen.SaveDocument(ownerID, UIUtilsGen.EntityTypes.Owner, strDOC_NAME, _
                                            IIf(isNonComplianceLetter, "Non Compliance Letter", "Compliance Letter"), DOC_PATH, _
                                            IIf(isNonComplianceLetter, "Non Compliance Letter", "Compliance Letter"), UIUtilsGen.ModuleID.Registration, 0, 0, 0)
                    oWord.Visible = True
                End If
                oWord = Nothing
            End If

        Catch ex As Exception
            Throw ex
        Finally
            SetCursorType(System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
#End Region
#Region "Closure Letters"
    Friend Function GenerateClosureLetter(ByVal EntityID As Integer, ByVal strDocType As String, ByVal strDocName As String, ByVal strDocDesc As String, ByVal strTemplateName As String, ByVal DueDate As Date, Optional ByRef pOwner As MUSTER.BusinessLogic.pOwner = Nothing, Optional ByVal AnalysisType As String = "", Optional ByVal PMHead As String = "", Optional ByVal nClosureID As Integer = 0, Optional ByVal strCertifiedContractor As String = "", Optional ByVal eventID As Int64 = 0, Optional ByVal eventSequence As Integer = 0, Optional ByVal eventType As Integer = 0, Optional ByVal media As String = " soil/ water") As Boolean

        Dim bolNeedCapDocs As Boolean = False
        Dim strDOC_NAME As String = String.Empty
        Dim strToday As String = String.Empty
        Dim nFacId As Integer = 0
        Dim oFacInfo As MUSTER.Info.FacilityInfo
        Dim oOwnerInfo As MUSTER.Info.OwnerInfo
        Dim oPersonaInfo As MUSTER.Info.PersonaInfo
        Dim oAddressInfo As MUSTER.Info.AddressInfo
        Dim colParams As New Specialized.NameValueCollection
        nModuleID = UIUtilsGen.ModuleID.Closure
        Try
            If DOC_PATH = "\" Then
                Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
            End If

            strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm")) + CStr(Format(Now, "ss"))
            strDOC_NAME = "CLO_" + strDocName.Trim.ToString + "_" + CStr(Trim(EntityID.ToString)) + "_" + strToday + ".doc"

            'Build NameValueCollection with Tags and Values.
            oOwnerInfo = pOwner.Retrieve(pOwner.ID, "SELF")
            oFacInfo = pOwner.Facilities.Retrieve(pOwner.OwnerInfo, EntityID, "SELF", "FACILITY")
            colParams.Add("<Title>", strDocType.ToString)
            colParams.Add("<DATE>", Format(Now, "MMMM d, yyyy"))


            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Contact name
            Dim dtContacts As DataTable
            Dim drRow As DataRow
            Dim strXHContactName As String = String.Empty
            Dim strXLContactName As String = String.Empty
            Dim strRXLContactName As String = String.Empty
            Dim Greeting As String = String.Empty

            dtContacts = GetXHAndXLContacts(nClosureID, UIUtilsGen.EntityTypes.ClosureEvent, UIUtilsGen.ModuleID.Closure, pOwner.ID, 1)
            If dtContacts.Rows.Count > 0 Then
                For Each drRow In dtContacts.Rows

                    Greeting = IIf(drRow("Greeting") Is DBNull.Value, "", drRow("Greeting"))

                    If drRow("EntityID") = nClosureID Then
                        If drRow("Type") = EnumContactType.XH Then
                            strXHContactName = drRow("CONTACT_Name")
                            colParams.Add("<Owner Address 1>", drRow("Address_One").ToString)
                            If drRow("Address_Two").ToString = String.Empty Then
                                colParams.Add("<Owner Address 2>", drRow("City").ToString & ", " & drRow("State").ToString & " " & drRow("Zip").ToString)
                                colParams.Add("<City/State/Zip>", "")
                            Else
                                colParams.Add("<Owner Address 2>", drRow("Address_Two").ToString)
                                colParams.Add("<City/State/Zip>", drRow("City").ToString & ", " & drRow("State").ToString & " " & drRow("Zip").ToString)
                            End If

                            If Greeting <> String.Empty Then
                                strXHContactName = Greeting
                            End If

                        ElseIf drRow("Type") = EnumContactType.XL Then
                            strXLContactName = drRow("CONTACT_Name")

                            If Greeting <> String.Empty Then
                                strXLContactName = Greeting
                            End If

                        End If
                    ElseIf drRow("EntityID") = pOwner.ID Then
                        If drRow("Type") = EnumContactType.XL Then
                            strRXLContactName = drRow("CONTACT_Name")

                            If Greeting <> String.Empty Then
                                strRXLContactName = Greeting
                            End If

                        End If
                    End If


                Next
                If strXHContactName <> String.Empty And strXLContactName <> String.Empty Then
                    colParams.Add("<Owner Name>", strXHContactName)
                    colParams.Add("<Contact Name>", strXLContactName)
                    colParams.Add("<Owner Greeting>", Greeting)
                ElseIf (strXHContactName = String.Empty And strXLContactName <> String.Empty) Then
                    colParams.Add("<Contact Name>", strXLContactName)
                    colParams.Add("<Owner Greeting>", Greeting)
                ElseIf strXHContactName <> String.Empty And strXLContactName = String.Empty Then
                    colParams.Add("<Owner Name>", strXHContactName)
                    colParams.Add("<Contact Name>", "")
                    colParams.Add("<Owner Greeting>", Greeting)
                ElseIf strRXLContactName <> String.Empty Then
                    colParams.Add("<Contact Name>", strRXLContactName)
                    colParams.Add("<Owner Greeting>", Greeting)
                End If
            End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'If (strXHContactName = String.Empty And strXLContactName = String.Empty) Or (strXHContactName = String.Empty) Then
            '    If oOwnerInfo.OrganizationID > 0 Then
            '        oPersonaInfo = pOwner.Organization
            '        colParams.Add("<Company Name>", pOwner.BPersona.Company)
            '        If strXLContactName = String.Empty And strRXLContactName = String.Empty Then
            '            colParams.Add("<Salutation>", pOwner.BPersona.Company.Trim)
            '        End If
            '        colParams.Add("<OWNER NAME>", pOwner.BPersona.Company)
            '    Else
            '        oPersonaInfo = pOwner.Persona
            '        colParams.Add("<Company Name>", pOwner.BPersona.FirstName.Trim & IIf(pOwner.BPersona.MiddleName.Trim.Length > 0, " ", "") & pOwner.BPersona.MiddleName.Trim & IIf(pOwner.BPersona.LastName.Trim.Length > 0, " ", "") & pOwner.BPersona.LastName.Trim & IIf(pOwner.BPersona.Suffix.Trim.Length > 0, " ", "") & pOwner.BPersona.Suffix.Trim)
            '        If strXLContactName = String.Empty And strRXLContactName = String.Empty Then
            '            colParams.Add("<Salutation>", "Dear " & pOwner.BPersona.Title.Trim & IIf(pOwner.BPersona.Title.Trim.Length > 0, " ", "") & pOwner.BPersona.FirstName.Trim & IIf(pOwner.BPersona.MiddleName.Trim.Length > 0, " ", "") & pOwner.BPersona.MiddleName.Trim & IIf(pOwner.BPersona.LastName.Trim.Length > 0, " ", "") & pOwner.BPersona.LastName.Trim & IIf(pOwner.BPersona.Suffix.Trim.Length > 0, " ", "") & pOwner.BPersona.Suffix.Trim & ":")
            '        End If
            '        'colParams.Add("<OWNER NAME>", pOwner.BPersona.FirstName.Trim & IIf(pOwner.BPersona.MiddleName.Trim.Length > 0, " ", "") & pOwner.BPersona.MiddleName.Trim & IIf(pOwner.BPersona.LastName.Trim.Length > 0, " ", "") & pOwner.BPersona.LastName.Trim & IIf(pOwner.BPersona.Suffix.Trim.Length > 0, " ", "") & pOwner.BPersona.Suffix.Trim)
            '    End If
            'End If
            'If dtContacts.Rows.Count = 0 Then
            '    colParams.Add("<Owner Contact>", "")
            'End If
            If (strXHContactName = String.Empty And strXLContactName = String.Empty) Or (strXHContactName = String.Empty) Then
                If oOwnerInfo.OrganizationID > 0 Then
                    oPersonaInfo = pOwner.Organization
                    colParams.Add("<Owner Name>", pOwner.BPersona.Company)
                    If strXLContactName = String.Empty And strRXLContactName = String.Empty Then
                        colParams.Add("<Owner Greeting>", pOwner.BPersona.Company.Trim)
                    End If
                Else
                    oPersonaInfo = pOwner.Persona
                    colParams.Add("<Owner Name>", pOwner.BPersona.Title.Trim & IIf(pOwner.BPersona.Title.Length > 0, " ", "") & pOwner.BPersona.FirstName.Trim & IIf(pOwner.BPersona.MiddleName.Trim.Length > 0, " ", "") & pOwner.BPersona.MiddleName.Trim & " " & pOwner.BPersona.LastName.Trim & " " & pOwner.BPersona.Suffix)
                    If strXLContactName = String.Empty And strRXLContactName = String.Empty Then
                        colParams.Add("<Owner Greeting>", pOwner.BPersona.FirstName.Trim & IIf(pOwner.BPersona.MiddleName.Trim.Length > 0, " ", "") & pOwner.BPersona.MiddleName.Trim & " " & pOwner.BPersona.LastName.Trim & IIf(pOwner.BPersona.Suffix.Trim.Length > 0, " ", "") & pOwner.BPersona.Suffix.Trim)
                    End If
                End If
            End If

            If dtContacts.Rows.Count = 0 Then
                colParams.Add("<Contact Name>", "")
            End If

            If strXHContactName = String.Empty Then
                oAddressInfo = pOwner.Address()
                colParams.Add("<Owner Address 1>", oAddressInfo.AddressLine1)
                If oAddressInfo.AddressLine2 = String.Empty Then
                    colParams.Add("<Owner Address 2>", oAddressInfo.City & ", " & oAddressInfo.State.TrimEnd & " " & oAddressInfo.Zip)
                    colParams.Add("<City/State/Zip>", "")

                Else
                    colParams.Add("<Owner Address 2>", oAddressInfo.AddressLine2)
                    colParams.Add("<City/State/Zip>", oAddressInfo.City & ", " & oAddressInfo.State.TrimEnd & " " & oAddressInfo.Zip & Chr(13) & Chr(9))

                End If
            End If

            oAddressInfo = pOwner.Facilities.FacilityAddresses.Retrieve(oFacInfo.AddressID)
            colParams.Add("<Facility Name>", oFacInfo.Name)
            colParams.Add("<Facility Address>", oAddressInfo.AddressLine1 & " " & IIf(oAddressInfo.AddressLine2 = String.Empty, "", oAddressInfo.AddressLine2))
            colParams.Add("<Facility City>", oAddressInfo.City.TrimEnd)
            colParams.Add("<I.D. #>", EntityID.ToString)
            colParams.Add("<Due Date>", DueDate.ToShortDateString)
            colParams.Add("<Schedule Date>", DueDate.ToShortDateString)
            colParams.Add("<User>", MusterContainer.AppUser.Name)
            colParams.Add("<User Phone>", CType(MusterContainer.AppUser.PhoneNumber, String))
            colParams.Add(" soil/ water", media)
            colParams.Add("<Soil/Water>", media)

            If AnalysisType = "BTEX" Then
                colParams.Add("<AnalysisType>", "Benzene, Toluene, Ethylbenzene, and Xylenes (BTEX)")
            ElseIf AnalysisType = "PAH" Then
                colParams.Add("<AnalysisType>", "Polynuclear Aromatic Hydrocarbons (PAH)")
            ElseIf AnalysisType = String.Empty Then
                colParams.Add("<AnalysisType>", "Unknown")
            Else
                colParams.Add("<AnalysisType>", "Benzene, Toluene, Ethylbenzene, and Xylenes (BTEX) and Polynuclear Aromatic Hydrocarbons (PAH)")
            End If

            colParams.Add("<PMHead>", PMHead.ToString)
            colParams.Add("<Certified Contractor>", strCertifiedContractor)
            FillCCList(nClosureID, UIUtilsGen.ModuleID.Closure, colParams)
            Try
                Dim strTempPath As String = TmpltPath & "Closure\" & strTemplateName
                Dim oWord As Word.Application = UIUtilsGen.CreateDocument("Closure", DOC_PATH, strDOC_NAME, strTempPath, colParams)
                If Not oWord Is Nothing Then
                    UIUtilsGen.SaveDocument(EntityID, 6, strDOC_NAME, strDocType, Doc_Path, strDocDesc, nModuleID, eventID, eventSequence, eventType)
                    oWord.Visible = True
                End If

            Catch ex As Exception
                Throw ex
            End Try

        Catch ex As Exception
            Throw ex
        Finally
            SetCursorType(System.Windows.Forms.Cursors.Default)
        End Try
    End Function
    Friend Function GenerateClosureInfoNeededLetter(ByVal facID As Integer, ByVal strDocType As String, ByVal strDocName As String, ByVal strDocDesc As String, ByVal strTemplateName As String, ByRef pOwner As MUSTER.BusinessLogic.pOwner, ByRef pClosure As MUSTER.BusinessLogic.pClosureEvent, ByVal colInfoNeeded As ArrayList, ByVal strCertifiedContractor As String) As Boolean
        Dim strDOC_NAME As String = String.Empty
        Dim strToday As String = String.Empty
        Dim colParams As New Specialized.NameValueCollection
        nModuleID = UIUtilsGen.ModuleID.Closure
        Try
            If DOC_PATH = "\" Then
                Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
            End If

            strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm")) + CStr(Format(Now, "ss"))
            strDOC_NAME = "CLO_" + strDocName.Trim.ToString + "_" + CStr(Trim(facID.ToString)) + "_" + strToday + ".doc"

            'Build NameValueCollection with Tags and Values.
            If pOwner.Facilities.ID <> facID Then
                pOwner.Facilities.Retrieve(pOwner.OwnerInfo, facID, "SELF", "FACILITY")
            End If
            colParams.Add("<Title>", strDocType.ToString)
            colParams.Add("<DATE>", Format(Now, "MMMM d, yyyy"))


            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Contact name
            Dim dtContacts As DataTable
            Dim drRow As DataRow
            Dim strXHContactName As String = String.Empty
            Dim strXLContactName As String = String.Empty
            Dim strRXLContactName As String = String.Empty
            Dim Greeting As String = String.Empty

            dtContacts = GetXHAndXLContacts(pClosure.ID, UIUtilsGen.EntityTypes.ClosureEvent, UIUtilsGen.ModuleID.Closure, pOwner.ID, 1)
            If dtContacts.Rows.Count > 0 Then
                For Each drRow In dtContacts.Rows
                    If drRow("EntityID") = pClosure.ID Then
                        If drRow("Type") = EnumContactType.XH Then
                            strXHContactName = drRow("CONTACT_Name")
                            colParams.Add("<Owner Address 1>", drRow("Address_One").ToString)
                            If drRow("Address_Two").ToString = String.Empty Then
                                colParams.Add("<Owner Address 2>", drRow("City").ToString & ", " & drRow("State").ToString & " " & drRow("Zip").ToString)
                                colParams.Add("<City/State/Zip>", "")
                            Else
                                colParams.Add("<Owner Address 2>", drRow("Address_Two").ToString)
                                colParams.Add("<City/State/Zip>", drRow("City").ToString & ", " & drRow("State").ToString & " " & drRow("Zip").ToString)
                            End If
                        ElseIf drRow("Type") = EnumContactType.XL Then
                            strXLContactName = drRow("CONTACT_Name")
                        End If
                    ElseIf drRow("EntityID") = pOwner.ID Then
                        If drRow("Type") = EnumContactType.XL Then
                            strRXLContactName = drRow("CONTACT_Name")
                        End If
                    End If
                    Greeting = drRow("Greeting")
                Next
                If strXHContactName <> String.Empty And strXLContactName <> String.Empty Then
                    colParams.Add("<Owner Name>", strXHContactName)
                    colParams.Add("<Contact Name>", strXLContactName)
                    colParams.Add("<Owner Greeting>", Greeting)
                ElseIf (strXHContactName = String.Empty And strXLContactName <> String.Empty) Then
                    colParams.Add("<Contact Name>", strXLContactName)
                    colParams.Add("<Owner Greeting>", Greeting)
                ElseIf strXHContactName <> String.Empty And strXLContactName = String.Empty Then
                    colParams.Add("<Owner Name>", strXHContactName)
                    colParams.Add("<Contact Name>", "")
                    colParams.Add("<Owner Greeting>", Greeting)
                ElseIf strRXLContactName <> String.Empty Then
                    colParams.Add("<Contact Name>", strRXLContactName)
                    colParams.Add("<Owner Greeting>", Greeting)
                End If
            End If

            If (strXHContactName = String.Empty And strXLContactName = String.Empty) Or (strXHContactName = String.Empty) Then
                If pOwner.OwnerInfo.OrganizationID > 0 Then
                    colParams.Add("<Owner Name>", pOwner.BPersona.Company)
                    If strXLContactName = String.Empty And strRXLContactName = String.Empty Then
                        colParams.Add("<Owner Greeting>", pOwner.BPersona.Company.Trim & ":")
                    End If
                Else
                    colParams.Add("<Owner Name>", pOwner.BPersona.Title.Trim & IIf(pOwner.BPersona.Title.Length > 0, " ", "") & pOwner.BPersona.FirstName.Trim & IIf(pOwner.BPersona.MiddleName.Trim.Length > 0, " ", "") & pOwner.BPersona.MiddleName.Trim & " " & pOwner.BPersona.LastName.Trim & " " & pOwner.BPersona.Suffix)
                    If strXLContactName = String.Empty And strRXLContactName = String.Empty Then
                        colParams.Add("<Owner Greeting>", pOwner.BPersona.Title.Trim & IIf(pOwner.BPersona.Title.Length > 0, " ", "") & pOwner.BPersona.FirstName.Trim & IIf(pOwner.BPersona.MiddleName.Trim.Length > 0, " ", "") & pOwner.BPersona.MiddleName.Trim & " " & pOwner.BPersona.LastName.Trim & " " & pOwner.BPersona.Suffix)
                    End If
                End If
            End If
            If dtContacts.Rows.Count = 0 Then
                colParams.Add("<Contact Name>", "")
            End If

            If strXHContactName = String.Empty Then
                colParams.Add("<Owner Address 1>", pOwner.Address.AddressLine1.Trim)
                If pOwner.Address.AddressLine2 = String.Empty Then
                    colParams.Add("<Owner Address 2>", pOwner.Address.City & ", " & pOwner.Address.State.TrimEnd & " " & pOwner.Address.Zip)
                    colParams.Add("<City/State/Zip>", "")
                Else
                    colParams.Add("<Owner Address 2>", pOwner.Address.AddressLine2)
                    colParams.Add("<City/State/Zip>", pOwner.Address.City & ", " & pOwner.Address.State.TrimEnd & " " & pOwner.Address.Zip)
                End If
            End If
            colParams.Add("<Facility Name>", pOwner.Facilities.Name)
            colParams.Add("<Facility Address>", pOwner.Facilities.FacilityAddresses.AddressLine1 & " " & IIf(pOwner.Facilities.FacilityAddresses.AddressLine2 = String.Empty, "", pOwner.Facilities.FacilityAddresses.AddressLine2))
            colParams.Add("<Facility City>", pOwner.Facilities.FacilityAddresses.City.TrimEnd)
            colParams.Add("<I.D. #>", facID.ToString)
            colParams.Add("<Due Date>", pClosure.DueDate.ToShortDateString)
            colParams.Add("<User>", MusterContainer.AppUser.Name)
            colParams.Add("<User Phone>", CType(MusterContainer.AppUser.PhoneNumber, String))
            colParams.Add("<Certified Contractor>", strCertifiedContractor)
            FillCCList(pClosure.ID, UIUtilsGen.ModuleID.Closure, colParams)

            Dim oWord As Word.Application = MusterContainer.GetWordApp

            If Not oWord Is Nothing Then

                Try
                    Dim strTempPath As String = TmpltPath & "Closure\" & strTemplateName
                    Dim ltrGen As New MUSTER.BusinessLogic.pLetterGen
                    ltrGen.CreateClosureInfoNeeded(strTempPath, DOC_PATH + strDOC_NAME, colParams, colInfoNeeded, oWord)
                    UIUtilsGen.SaveDocument(facID, UIUtilsGen.EntityTypes.Facility, strDOC_NAME, strDocType, Doc_Path, strDocDesc, nModuleID, pClosure.ID, pClosure.FacilitySequence, UIUtilsGen.EntityTypes.ClosureEvent)
                    'UIUtilsGen.CreateAndSaveDocument("Closure", facID, uiutilsgen.EntityTypes.Facility, DOC_PATH, strDOC_NAME, strDocType, strTempPath, Doc_Path, strDocDesc, colParams)
                    oWord.Visible = True
                Catch ex As Exception
                    Throw ex
                End Try
            End If
            oWord = Nothing



        Catch ex As Exception
            Throw ex
        Finally
            SetCursorType(System.Windows.Forms.Cursors.Default)
        End Try
    End Function
    Friend Function GenerateSampleDemo(ByVal EntityID As Integer, ByVal strDocType As String, ByVal strDocName As String, ByVal strDocDesc As String, ByVal strTemplateName As String, ByVal DueDate As Date, Optional ByRef pOwner As MUSTER.BusinessLogic.pOwner = Nothing, Optional ByVal strPMHead As String = "", Optional ByVal strClosureHead As String = "", Optional ByVal dtSample As DataTable = Nothing, Optional ByVal nClosureID As Integer = 0, Optional ByVal strCertifiedContractor As String = "", Optional ByVal eventID As Int64 = 0, Optional ByVal eventSequence As Integer = 0, Optional ByVal eventType As Integer = 0) As Boolean

        Dim bolNeedCapDocs As Boolean = False
        Dim strDOC_NAME As String = String.Empty
        Dim strToday As String = String.Empty
        Dim nFacId As Integer = 0
        Dim oFacInfo As MUSTER.Info.FacilityInfo
        Dim oOwnerInfo As MUSTER.Info.OwnerInfo
        Dim oPersonaInfo As MUSTER.Info.PersonaInfo
        Dim oAddressInfo As MUSTER.Info.AddressInfo
        Dim colParams As New Specialized.NameValueCollection
        Dim strAddress1 As String
        Dim strAddress2 As String
        Dim strCity As String
        Dim strState As String
        Dim strZip As String
        Dim strPhone As String
        nModuleID = UIUtilsGen.ModuleID.Closure
        Try

            If DOC_PATH = "\" Then
                Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
            End If

            strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm")) + CStr(Format(Now, "ss"))
            strDOC_NAME = "CLO_" + strDocName.Trim.ToString + "_" + CStr(Trim(EntityID.ToString)) + "_" + strToday + ".doc"

            'Build NameValueCollection with Tags and Values.
            oOwnerInfo = pOwner.Retrieve(pOwner.ID, "SELF")
            colParams.Add("<Title>", strDocType.ToString)
            colParams.Add("<DateGenerated>", Format(Now, "MMMM d, yyyy"))
            colParams.Add("<PMHead>", strPMHead)
            colParams.Add("<ClosureHead>", strClosureHead)


            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Contact name
            Dim dtContacts As DataTable
            Dim drRow As DataRow
            Dim strXHContactName As String = String.Empty
            Dim strXLContactName As String = String.Empty
            Dim strRXLContactName As String = String.Empty

            dtContacts = GetXHAndXLContacts(nClosureID, UIUtilsGen.EntityTypes.ClosureEvent, UIUtilsGen.ModuleID.Closure, pOwner.ID, 1)
            If dtContacts.Rows.Count > 0 Then
                For Each drRow In dtContacts.Rows
                    If drRow("EntityID") = nClosureID Then
                        If drRow("Type") = EnumContactType.XH Then
                            strXHContactName = drRow("Greeting")
                            strAddress1 = drRow("Address_One").ToString
                            If drRow("Address_Two").ToString = String.Empty Then
                                strAddress2 = ""
                            Else
                                strAddress2 = drRow("Address_Two").ToString
                            End If
                            strCity = drRow("City").ToString
                            strState = drRow("State").ToString
                            strZip = drRow("Zip").ToString
                            strPhone = drRow("Phone").ToString

                        ElseIf drRow("Type") = EnumContactType.XL Then
                            strXLContactName = drRow("Greeting")
                        End If
                    ElseIf drRow("EntityID") = pOwner.ID Then
                        If drRow("Type") = EnumContactType.XL Then
                            strRXLContactName = drRow("Greeting")
                        End If
                    End If
                Next
                If strXHContactName <> String.Empty And strXLContactName <> String.Empty Then
                    colParams.Add("<Owner Name>", strXHContactName)
                    colParams.Add("<Contact Name>", strXLContactName)
                    colParams.Add("<Owner Address 1>", strAddress1)
                    If strAddress2 = String.Empty Then
                        colParams.Add("<Owner Address 2>", strCity & ", " & strState & " " & strZip)
                        colParams.Add("<City/State/Zip>", strPhone)
                        colParams.Add("<Owner Phone>", "")
                    Else
                        colParams.Add("<Owner Address 2>", strAddress2)
                        colParams.Add("<Owner City/State/Zip>", strCity & ", " & strState & " " & strZip)
                        colParams.Add("<Owner Phone>", strPhone)
                    End If

                ElseIf (strXHContactName = String.Empty And strXLContactName <> String.Empty) Then
                    colParams.Add("<Contact Name>", strXLContactName)
                ElseIf strXHContactName <> String.Empty And strXLContactName = String.Empty Then
                    colParams.Add("<Contact Name>", strXHContactName)
                    colParams.Add("<Owner Name>", strAddress1)
                    If strAddress2 = String.Empty Then
                        colParams.Add("<Owner Address 1>", strCity & ", " & strState & " " & strZip)
                        colParams.Add("<Owner Address 2>", strPhone)
                        colParams.Add("<City/State/Zip>", "")
                        colParams.Add("<Owner Phone>", "")
                    Else
                        colParams.Add("<Owner Address 1>", strAddress2)
                        colParams.Add("<Owner Address 2>", strCity & ", " & strState & " " & strZip)
                        colParams.Add("<City/State/Zip>", strPhone)
                        colParams.Add("<Owner Phone>", "")
                    End If

                ElseIf strRXLContactName <> String.Empty Then
                    colParams.Add("<Contact Name>", strRXLContactName)
                End If
            End If

            If (strXLContactName = String.Empty And strRXLContactName = String.Empty) Then

                If oOwnerInfo.OrganizationID > 0 Then
                    oPersonaInfo = pOwner.Organization
                    colParams.Add("<Contact Name>", pOwner.BPersona.Company)
                Else
                    oPersonaInfo = pOwner.Persona
                    colParams.Add("<Contact Name>", pOwner.BPersona.Title.Trim & IIf(pOwner.BPersona.Title.Length > 0, " ", "") & pOwner.BPersona.FirstName.Trim & IIf(pOwner.BPersona.MiddleName.Trim.Length > 0, " ", "") & pOwner.BPersona.MiddleName.Trim & " " & pOwner.BPersona.LastName.Trim & " " & pOwner.BPersona.Suffix)
                End If
            Else
                If (strXHContactName = String.Empty And strXLContactName = String.Empty) Or (strXHContactName = String.Empty) Then
                    If oOwnerInfo.OrganizationID > 0 Then
                        oPersonaInfo = pOwner.Organization
                        colParams.Add("<Owner Name>", pOwner.BPersona.Company)
                    Else
                        oPersonaInfo = pOwner.Persona
                        colParams.Add("<Owner Name>", pOwner.BPersona.Title.Trim & IIf(pOwner.BPersona.Title.Length > 0, " ", "") & pOwner.BPersona.FirstName.Trim & IIf(pOwner.BPersona.MiddleName.Trim.Length > 0, " ", "") & pOwner.BPersona.MiddleName.Trim & " " & pOwner.BPersona.LastName.Trim & " " & pOwner.BPersona.Suffix)
                    End If
                End If
            End If
            If (strXLContactName = String.Empty And strRXLContactName = String.Empty) Then
                If strXHContactName = String.Empty Then
                    oAddressInfo = pOwner.Address()
                    colParams.Add("<Owner Name>", oAddressInfo.AddressLine1)
                    If oAddressInfo.AddressLine2 = String.Empty Then
                        colParams.Add("<Owner Address 1>", oAddressInfo.City & ", " & oAddressInfo.State.TrimEnd & " " & oAddressInfo.Zip)
                        colParams.Add("<Owner Address 2>", pOwner.PhoneNumberOne)
                        colParams.Add("<City/State/Zip>", "")
                        colParams.Add("<Owner Phone>", "")
                    Else
                        colParams.Add("<Owner Address 1>", oAddressInfo.AddressLine2)
                        colParams.Add("<Owner Address 2>", oAddressInfo.City & ", " & oAddressInfo.State.TrimEnd & " " & oAddressInfo.Zip)
                        colParams.Add("<City/State/Zip>", pOwner.PhoneNumberOne)
                        colParams.Add("<Owner Phone>", "")
                    End If
                End If
                'colParams.Add("<Owner Phone>", pOwner.PhoneNumberOne)
            Else
                If strXHContactName = String.Empty Then
                    oAddressInfo = pOwner.Address()
                    colParams.Add("<Owner Address 1>", oAddressInfo.AddressLine1)
                    If oAddressInfo.AddressLine2 = String.Empty Then
                        colParams.Add("<Owner Address 2>", oAddressInfo.City & ", " & oAddressInfo.State.TrimEnd & " " & oAddressInfo.Zip)
                        colParams.Add("<City/State/Zip>", pOwner.PhoneNumberOne)
                        colParams.Add("<Owner Phone>", "")
                    Else
                        colParams.Add("<Owner Address 2>", oAddressInfo.AddressLine2)
                        colParams.Add("<Owner City/State/Zip>", oAddressInfo.City & ", " & oAddressInfo.State.TrimEnd & " " & oAddressInfo.Zip)
                        colParams.Add("<Owner Phone>", pOwner.PhoneNumberOne)
                    End If

                End If


            End If


            'oAddressInfo = pOwner.Addresses.Retrieve(oFacInfo.AddressID)
            oFacInfo = pOwner.Facilities.Retrieve(pOwner.OwnerInfo, EntityID, "SELF", "FACILITY")
            Dim pAddress As New MUSTER.BusinessLogic.pAddress
            pAddress.Retrieve(oFacInfo.AddressID)
            'Dim AddrForm As New GenAddressMSFT
            'AddrForm.GetAddress(oFacInfo.AddressID)
            'AddrForm.ShowFIPS = False
            'AddrForm.ShowCounty = True

            colParams.Add("<Facility Name>", oFacInfo.Name)
            colParams.Add("<Facility Address>", pAddress.AddressLine1 + IIf(pAddress.AddressLine2 = String.Empty, String.Empty, pAddress.AddressLine2))
            colParams.Add("<Facility City>", pAddress.City.Trim)
            colParams.Add("<Facility County>", pAddress.County)
            colParams.Add("<I.D. #>", EntityID.ToString)
            colParams.Add("<Certified Contractor>", strCertifiedContractor)
            FillCCList(nClosureID, UIUtilsGen.ModuleID.Closure, colParams)
            'AddrForm = Nothing
            pAddress = Nothing

            Dim oWord As Word.Application = MusterContainer.GetWordApp

            If Not oWord Is Nothing Then


                Try

                    Dim strTempPath As String = TmpltPath & "Closure\" & strTemplateName
                    Dim ltrGen As New MUSTER.BusinessLogic.pLetterGen
                    ltrGen.CreateClosureDemo("Closure", strDocName, colParams, strTempPath, Doc_Path & strDOC_NAME, dtSample, oWord)
                    UIUtilsGen.SaveDocument(EntityID, 6, strDOC_NAME, strDocType, Doc_Path, strDocDesc, nModuleID, eventID, eventSequence, eventType)
                    oWord.Visible = True
                Catch ex As Exception
                    Throw ex
                End Try
            End If
            oWord = Nothing



        Catch ex As Exception
            Throw ex
        Finally
            SetCursorType(System.Windows.Forms.Cursors.Default)
        End Try
    End Function
#End Region
#Region "Technical Letters"
    Friend Function GenerateTFCheckList(ByVal LustEventID As Integer, ByVal FacID As Integer, Optional ByRef pOwner As MUSTER.BusinessLogic.pOwner = Nothing, Optional ByVal bolDraft As Boolean = False, Optional ByVal eventSequence As Integer = 0) As Boolean

        Dim bolNeedCapDocs As Boolean = False
        Dim strDOC_NAME As String = String.Empty
        Dim strToday As String = String.Empty
        Dim nFacId As Integer = 0
        Dim oFacInfo As MUSTER.Info.FacilityInfo
        Dim oOwnerInfo As MUSTER.Info.OwnerInfo
        Dim oPersonaInfo As MUSTER.Info.PersonaInfo
        Dim oAddressInfo As MUSTER.Info.AddressInfo
        Dim oUserInfo As New MUSTER.BusinessLogic.pUser
        Dim oLustevent As New MUSTER.BusinessLogic.pLustEvent
        Dim aryTF As Array
        Dim strTFChecklist As String
        Dim i As Int16
        Dim strYes As String
        Dim strNo As String
        Dim strOther As String
        Dim colParams As New Specialized.NameValueCollection
        Dim strDocPath As String
        Dim tmpDate As Date

        nModuleID = UIUtilsGen.ModuleID.Technical

        Try
            oLustevent.Retrieve(LustEventID)
            oUserInfo.Retrieve(oLustevent.PM)

            If DOC_PATH = "\" Then
                Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
            End If

            strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm")) + "_" + CStr(Format(Now, "ss"))
            strDOC_NAME = "FAC_TFE_" + CStr(Trim(FacID.ToString)) + "_" + strToday + ".doc"

            oOwnerInfo = pOwner.Retrieve(pOwner.ID, "SELF")
            colParams.Add("<PM>", oUserInfo.Name)

            If oOwnerInfo.OrganizationID > 0 Then
                oPersonaInfo = pOwner.Organization
                colParams.Add("<Tank_Owner>", pOwner.BPersona.Company)
            Else
                oPersonaInfo = pOwner.Persona
                colParams.Add("<Tank_Owner>", pOwner.BPersona.Title.Trim & IIf(pOwner.BPersona.Title.Length > 0, " ", "") & pOwner.BPersona.FirstName.Trim & IIf(pOwner.BPersona.MiddleName.Trim.Length > 0, " ", "") & pOwner.BPersona.MiddleName.Trim & " " & pOwner.BPersona.LastName.Trim & " " & pOwner.BPersona.Suffix)
            End If

            oAddressInfo = pOwner.Address()
            colParams.Add("<TO_Address1>", oAddressInfo.AddressLine1)
            colParams.Add("<TO_Address2>", oAddressInfo.City & ", " & oAddressInfo.State.TrimEnd & " " & oAddressInfo.Zip)
            colParams.Add("<Report_Date>", IIf(oLustevent.ReportDate = tmpDate, "", oLustevent.ReportDate))
            colParams.Add("<Start_Date>", IIf(oLustevent.Started = tmpDate, "", oLustevent.Started))
            colParams.Add("<Facility_ID>", oLustevent.FacilityID)
            colParams.Add("<Comments>", oLustevent.ELIGIBITY_COMMENTS)


            If oLustevent.TFCheckList Is Nothing Then
                strTFChecklist = "X|X|X|X|X|X|X|X|X|X|X|X|X|X|X|X|X"
            Else
                strTFChecklist = oLustevent.TFCheckList
            End If

            If strTFChecklist.Length < 1 Then
                strTFChecklist = "X|X|X|X|X|X|X|X|X|X|X|X|X|X|X|X|X"
            End If
            aryTF = Split(strTFChecklist, "|")


            For i = 0 To 16
                strYes = " "
                strNo = " "
                strOther = " "

                Select Case aryTF(i)
                    Case "Y"
                        strYes = "X"
                    Case "N"
                        strNo = "X"
                    Case "A"
                        strOther = "X"
                End Select
                colParams.Add("<" & CStr(i + 1) & "Y>", strYes)
                colParams.Add("<" & CStr(i + 1) & "N>", strNo)
                colParams.Add("<" & CStr(i + 1) & "U>", strOther)
            Next

            strYes = " "
            strNo = " "
            strOther = " "

            Select Case oLustevent.PM_HEAD_ASSESS
                Case 1                          ' Yes
                    strYes = "X"
                Case 3                          ' Undecided
                    strOther = "X"
            End Select
            colParams.Add("<PY>", strYes)
            colParams.Add("<PU>", strOther)

            strYes = " "
            strOther = " "

            Select Case oLustevent.UST_CHIEF_ASSESS
                Case 1                          ' Yes
                    strYes = "X"
                Case 2                          ' No
                    strNo = "X"
                Case 3                          ' Undecided
                    strOther = "X"
            End Select
            colParams.Add("<UY>", strYes)
            colParams.Add("<UN>", strNo)
            colParams.Add("<UU>", strOther)

            strYes = " "
            strNo = " "
            strOther = " "

            Select Case oLustevent.OPC_HEAD_ASSESS
                Case 1                          ' Yes
                    strYes = "X"
                Case 2                          ' No
                    strNo = "X"
            End Select
            colParams.Add("<OY>", strYes)
            colParams.Add("<ON>", strNo)

            Try
                Dim ltrGen As New MUSTER.BusinessLogic.pLetterGen
                Dim strTempPath As String = System.IO.Path.GetTempPath


                If bolDraft Then
                    strDocPath = strTempPath
                Else
                    strDocPath = Doc_Path
                End If

                Dim oWordApp As Word.Application = MusterContainer.GetWordApp

                If Not oWordApp Is Nothing Then


                    ltrGen.CreateLetter("Technical", strDOC_NAME, colParams, TmpltPath & "Technical\TF_Checklist_Template1.doc", strDocPath & strDOC_NAME, oWordApp)
                    'Delay()
                    UIUtilsGen.Delay(, 1)
                    Dim fileName As Object
                    Dim aDoc As Word.Document
                    Dim areadOnly As Object = True
                    Dim isVisible As Object = True
                    Dim confirmConversions As Object = False
                    Dim addToRecentFiles As Object = False
                    Dim revert As Object = False
                    ' Here is the way to handle parameters you don't care about in .NET
                    Dim missing As Object = System.Reflection.Missing.Value

                    ' Make word visible
                    oWordApp.Visible = True

                    If bolDraft Then


                        ' Open the document that was chosen by the dialog
                        aDoc = oWordApp.Documents.Open(strDocPath & strDOC_NAME, confirmConversions, areadOnly, addToRecentFiles, missing, missing, revert, missing, missing, missing, missing, isVisible)

                    Else
                        'UIUtilsGen.SaveDocument(oLustevent.ID, 7, strDOC_NAME, "TFE Checklist", DOC_PATH, "Trust Fund Eligibility Checklist", nModuleID)
                        UIUtilsGen.SaveDocument(FacID, UIUtilsGen.EntityTypes.Facility, strDOC_NAME, "TFE Checklist", DOC_PATH, "Trust Fund Eligibility Checklist", nModuleID, oLustevent.ID, eventSequence, UIUtilsGen.EntityTypes.LUST_Event)
                    End If
                End If
                oWordApp = Nothing

            Catch ex As Exception
                Throw ex
            End Try

        Catch ex As Exception
            Throw ex
        Finally
            SetCursorType(System.Windows.Forms.Cursors.Default)
        End Try
    End Function
    Friend Function GenerateTechLetter(ByVal FacilityID As Integer, ByVal strDocType As String, ByVal strDocName As String, ByVal strDocDesc As String, ByVal strTemplateName As String, ByVal DueDate As Date, ByVal EventID As Int64, Optional ByRef pOwner As MUSTER.BusinessLogic.pOwner = Nothing, Optional ByVal nCertifiedMailNo As Integer = 0, Optional ByVal eventSequence As Integer = 0, Optional ByVal eventType As Integer = 0) As Boolean
        Dim strContactName As String
        Dim strOwnerName As String
        Dim bolNeedCapDocs As Boolean = False
        Dim strDOC_NAME As String = String.Empty
        Dim strToday As String = String.Empty
        Dim nFacId As Integer = 0

        Dim oFacInfo As MUSTER.Info.FacilityInfo
        Dim oOwnerInfo As MUSTER.Info.OwnerInfo
        Dim oPersonaInfo As MUSTER.Info.PersonaInfo
        Dim oAddressInfo As MUSTER.Info.AddressInfo

        Dim oUserPM As New MUSTER.BusinessLogic.pUser
        Dim oUserSupervisor As New MUSTER.BusinessLogic.pUser
        Dim oLustEvent As New MUSTER.BusinessLogic.pLustEvent
        Dim oERACContact As New MUSTER.BusinessLogic.pCompany
        Dim oIRACContact As New MUSTER.BusinessLogic.pCompany
        Dim oContact As New MUSTER.BusinessLogic.pContactDatum
        Dim pContact As New MUSTER.BusinessLogic.pContactStruct
        Dim oContactInfo As New MUSTER.Info.ContactDatumInfo

        Dim colParams As New Specialized.NameValueCollection
        nModuleID = UIUtilsGen.ModuleID.Technical
        Try
            If DOC_PATH = "\" Then
                Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
            End If
            strContactName = ""

            strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm"))
            strDOC_NAME = "TECH_" + strDocName.Trim.ToString + "_" + CStr(Trim(FacilityID.ToString)) + "_" + strToday + ".doc"

            'Build NameValueCollection with Tags and Values.
            oOwnerInfo = pOwner.Retrieve(pOwner.ID, "SELF")
            oFacInfo = pOwner.Facilities.Retrieve(pOwner.OwnerInfo, FacilityID, "SELF", "FACILITY")
            If oOwnerInfo.ID <= 0 Then
                oOwnerInfo = pOwner.Retrieve(oFacInfo.OwnerID, "SELF")
            End If
            colParams.Add("<Title>", strDocType.ToString)
            colParams.Add("<DATE>", Format(Now, "MMMM d, yyyy"))
            oLustEvent.Retrieve(EventID)
            oUserPM.Retrieve(oLustEvent.PM)



            If nCertifiedMailNo <> 0 Then
                colParams.Add("<Certified Number>", nCertifiedMailNo.ToString)
            End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Contact name
            Dim dtContacts As DataTable
            Dim drRow As DataRow
            Dim strXHContactName As String = String.Empty
            Dim strXLContactName As String = String.Empty
            Dim strRXLContactName As String = String.Empty

            dtContacts = GetXHAndXLContacts(EventID, UIUtilsGen.EntityTypes.LUST_Event, UIUtilsGen.ModuleID.Technical, pOwner.ID, 1)
            If dtContacts.Rows.Count > 0 Then
                For Each drRow In dtContacts.Rows
                    If drRow("EntityID") = EventID Then
                        If drRow("Type") = EnumContactType.XH Then
                            strXHContactName = drRow("CONTACT_Name")
                            colParams.Add("<Company Address>", drRow("Address_One").ToString & " " & IIf(drRow("Address_Two").ToString = String.Empty, "", "; " & drRow("Address_Two").ToString))
                            colParams.Add("<City, State, Zip>", drRow("City").ToString & ", " & drRow("State").ToString & " " & drRow("Zip").ToString)
                        ElseIf drRow("Type") = EnumContactType.XL Then
                            strXLContactName = drRow("CONTACT_Name")
                        End If
                    ElseIf drRow("EntityID") = pOwner.ID Then
                        If drRow("Type") = EnumContactType.XL Then
                            strRXLContactName = drRow("CONTACT_Name")
                        End If
                    End If
                Next
                If strXHContactName <> String.Empty And strXLContactName <> String.Empty Then
                    colParams.Add("<Company Name>", strXHContactName)
                    colParams.Add("<Owner Contact>", strXLContactName)
                    colParams.Add("<Salutation>", strXLContactName)
                ElseIf (strXHContactName = String.Empty And strXLContactName <> String.Empty) Then
                    colParams.Add("<Owner Contact>", strXLContactName)
                    colParams.Add("<Salutation>", strXLContactName)
                ElseIf strXHContactName <> String.Empty And strXLContactName = String.Empty Then
                    colParams.Add("<Company Name>", strXHContactName)
                    colParams.Add("<Owner Contact>", "")
                    colParams.Add("<Salutation>", strXHContactName)
                ElseIf strRXLContactName <> String.Empty Then
                    colParams.Add("<Owner Contact>", strRXLContactName)
                    colParams.Add("<Salutation>", strRXLContactName)
                End If
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If (strXHContactName = String.Empty And strXLContactName = String.Empty) Or (strXHContactName = String.Empty) Then
                If oOwnerInfo.OrganizationID > 0 Then
                    oPersonaInfo = pOwner.Organization
                    colParams.Add("<Company Name>", pOwner.BPersona.Company)
                    If strXLContactName = String.Empty And strRXLContactName = String.Empty Then
                        colParams.Add("<Salutation>", pOwner.BPersona.Company.Trim)
                    End If
                    colParams.Add("<OWNER NAME>", pOwner.BPersona.Company)
                Else
                    oPersonaInfo = pOwner.Persona
                    colParams.Add("<Company Name>", pOwner.BPersona.FirstName.Trim & IIf(pOwner.BPersona.MiddleName.Trim.Length > 0, " ", "") & pOwner.BPersona.MiddleName.Trim & IIf(pOwner.BPersona.LastName.Trim.Length > 0, " ", "") & pOwner.BPersona.LastName.Trim & IIf(pOwner.BPersona.Suffix.Trim.Length > 0, " ", "") & pOwner.BPersona.Suffix.Trim)
                    If strXLContactName = String.Empty And strRXLContactName = String.Empty Then
                        colParams.Add("<Salutation>", pOwner.BPersona.FirstName.Trim & IIf(pOwner.BPersona.MiddleName.Trim.Length > 0, " ", "") & pOwner.BPersona.MiddleName.Trim & IIf(pOwner.BPersona.LastName.Trim.Length > 0, " ", "") & pOwner.BPersona.LastName.Trim & IIf(pOwner.BPersona.Suffix.Trim.Length > 0, " ", "") & pOwner.BPersona.Suffix.Trim & ":")
                    End If
                    'colParams.Add("<OWNER NAME>", pOwner.BPersona.FirstName.Trim & IIf(pOwner.BPersona.MiddleName.Trim.Length > 0, " ", "") & pOwner.BPersona.MiddleName.Trim & IIf(pOwner.BPersona.LastName.Trim.Length > 0, " ", "") & pOwner.BPersona.LastName.Trim & IIf(pOwner.BPersona.Suffix.Trim.Length > 0, " ", "") & pOwner.BPersona.Suffix.Trim)
                End If
            End If
            If dtContacts.Rows.Count = 0 Then
                colParams.Add("<Owner Contact>", "")
            End If
            'oContact.GetAllByEntity(oOwnerInfo.ID, 614)

            'If oContact.ContactCollection.Count > 0 Then
            '    For Each oContactInfo In oContact.ContactCollection.Values
            '        If oContactInfo.orgCode > 0 Then
            '            strContactName = oContactInfo.companyName
            '        Else
            '            strContactName = oContactInfo.fullName
            '        End If
            '    Next
            'Else
            '    'If oOwnerInfo.OrganizationID > 0 Then
            '    '    strContactName = pOwner.BPersona.Company
            '    'Else
            '    '    strContactName = pOwner.BPersona.FirstName.Trim & IIf(pOwner.BPersona.MiddleName.Trim.Length > 0, " ", "") & pOwner.BPersona.MiddleName.Trim & IIf(pOwner.BPersona.LastName.Trim.Length > 0, " ", "") & pOwner.BPersona.LastName.Trim & IIf(pOwner.BPersona.Suffix.Trim.Length > 0, " ", "") & pOwner.BPersona.Suffix.Trim
            '    'End If
            'End If
            '''''''''''''''''''''''''''''''''''''''''


            If strXHContactName = String.Empty Then
                oAddressInfo = pOwner.Address()
                'colParams.Add("<Company Address>", oAddressInfo.AddressLine1)
                colParams.Add("<Company Address>", oAddressInfo.AddressLine1 & " " & IIf(oAddressInfo.AddressLine2 = String.Empty, "", "; " & oAddressInfo.AddressLine2))
                colParams.Add("<City, State, Zip>", oAddressInfo.City & ", " & oAddressInfo.State.TrimEnd & " " & oAddressInfo.Zip)
            End If
            oAddressInfo = pOwner.Addresses.Retrieve(oFacInfo.AddressID)
            colParams.Add("<Facility Name>", oFacInfo.Name)
            colParams.Add("<Facility Address>", oAddressInfo.AddressLine1 & " " & IIf(oAddressInfo.AddressLine2 = String.Empty, "", "; " & oAddressInfo.AddressLine2))
            colParams.Add("<Facility City>", oAddressInfo.City.TrimEnd)
            colParams.Add("<I.D. #>", oFacInfo.ID.ToString)
            colParams.Add("<Due Date>", DueDate.ToShortDateString)
            colParams.Add("<Letter Date>", Today.ToShortDateString)

            If IsNothing(oUserPM.Name) Then
                colParams.Add("<Project Manager>", "")
            Else
                colParams.Add("<Project Manager>", oUserPM.Name)
            End If
            If IsNothing(oUserPM.PhoneNumber) Then
                colParams.Add("<PM Phone Number>", "")
            Else
                colParams.Add("<PM Phone Number>", oUserPM.PhoneNumber)
            End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'CC List
            FillCCList(EventID, UIUtilsGen.ModuleID.Technical, colParams)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If oUserPM.ManagerID > 0 Then
                oUserSupervisor.Retrieve(oUserPM.ManagerID)
                'colParams.Add("<CC List>", IIf(oUserSupervisor.Name Is Nothing, String.Empty, "CC :" & oUserSupervisor.Name & ", OPC"))
                colParams.Add("<CC Supervisor>", IIf(oUserSupervisor.Name Is Nothing, String.Empty, "" & oUserSupervisor.Name & ", OPC"))
                colParams.Add("<Supervisor Not Martha>", IIf(oUserSupervisor.Name Is Nothing Or oUserSupervisor.Name = "Martha Martin", String.Empty, "" & oUserSupervisor.Name & ", OPC"))
                colParams.Add("<Supervisor>", IIf(oUserSupervisor.Name Is Nothing, String.Empty, "" & oUserSupervisor.Name & ", OPC"))
            Else
                'colParams.Add("<CC List>", String.Empty)
                colParams.Add("<CC Supervisor>", String.Empty)
                colParams.Add("<Supervisor Not Martha>", String.Empty)
                colParams.Add("<Supervisor>", String.Empty)
            End If

            Dim strEracIrac As String = String.Empty
            If oLustEvent.ERAC > 0 Or oLustEvent.ERAC < -100 Then
                oERACContact.Retrieve(oLustEvent.ERAC)
                ' need erac rep name
                strEracIrac = GetEracIracRep(EventID, UIUtilsGen.ModuleID.Technical, "ERAC")
                strEracIrac += oERACContact.COMPANY_NAME
                colParams.Add("<ERAC>", strEracIrac)
                colParams.Add("<ERACCompany>", oERACContact.COMPANY_NAME)
                'Erac name & Address.
                Dim oComAdd As New MUSTER.BusinessLogic.pComAddress
                Dim dsComAdd As DataSet
                Dim drow As DataRow
                dsComAdd = oComAdd.GetCompanyAddress(oERACContact.ID)
                If dsComAdd.Tables(0).Rows.Count > 0 Then
                    For Each drow In dsComAdd.Tables(0).Rows
                        If Not drow Is Nothing Then
                            colParams.Add("<ERAC Address>", IIf(drow("ADDRESS_LINE_ONE") Is DBNull.Value, "", "" & drow("ADDRESS_LINE_ONE")))
                            colParams.Add("<ERAC City, State, Zip>", IIf(drow("CITY") Is DBNull.Value, "", drow("CITY")) & ", " & IIf(drow("STATE") = String.Empty, "", drow("STATE")) & " " & IIf(drow("ZIP") = String.Empty, "", drow("ZIP")))
                            Exit For
                        End If
                    Next
                End If
            Else
                colParams.Add("<ERAC>", strEracIrac)
                colParams.Add("<ERACCompany>", "")
                colParams.Add("<ERAC Address>", "")
                colParams.Add("<ERAC City, State, Zip>", "")
            End If

            strEracIrac = String.Empty
            If oLustEvent.IRAC > 0 Then
                oIRACContact.Retrieve(oLustEvent.IRAC)
                ' need erac rep name
                strEracIrac = GetEracIracRep(EventID, UIUtilsGen.ModuleID.Technical, "IRAC")
                strEracIrac += oIRACContact.COMPANY_NAME
                colParams.Add("<IRAC>", strEracIrac)
                colParams.Add("<IRACCompany>", oIRACContact.COMPANY_NAME)
            Else
                colParams.Add("<IRAC>", strEracIrac)
                colParams.Add("<IRACCompany>", "")
            End If



            '<EXTENSION DUE DATE>


            Try
                Dim strTempPath As String = TmpltPath & "Technical\" & strTemplateName
                UIUtilsGen.CreateAndSaveDocument("Technical", FacilityID, UIUtilsGen.EntityTypes.Facility, DOC_PATH, strDOC_NAME, strDocType, strTempPath, Doc_Path, strDocDesc, nModuleID, colParams, EventID, eventSequence, eventType)
            Catch ex As Exception
                Throw ex
            End Try

        Catch ex As Exception
            Throw ex
        Finally
            SetCursorType(System.Windows.Forms.Cursors.Default)
        End Try
    End Function
#End Region
#Region "Financial Letters"
    Friend Function GenerateFinancialUnencumberanceMemoTemplate(ByVal dtTable As DataTable, ByVal strDocType As String, ByVal strDocName As String, ByVal strDocDesc As String, ByVal strTemplateName As String, Optional ByRef pOwner As MUSTER.BusinessLogic.pOwner = Nothing, Optional ByVal nOwnerID As Integer = 0) As Boolean
        Dim strDOC_NAME As String = String.Empty
        Dim DestDoc As Word.Document
        Dim strToday As String = String.Empty
        Dim bolBreak As Boolean = False
        Dim nEntityID As Int16
        Dim nOwningEntity As Int64
        Dim dr As DataRow
        nModuleID = UIUtilsGen.ModuleID.Financial
        Try
            If DOC_PATH = "\" Then
                Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
            End If


            strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm")) + CStr(Format(Now, "ss"))
            strDOC_NAME = "FIN_" + strDocName.Trim.ToString + "_" + strToday + ".doc"
            Dim strTempPath As String = TmpltPath & "Financial\" & strTemplateName
            If Not System.IO.File.Exists(strTempPath) Then
                Throw New Exception("File Not Found: " + strTempPath)
            End If
            System.IO.File.Copy(strTempPath, Doc_Path & strDOC_NAME)


            Dim oWord As Word.Application = MusterContainer.GetWordApp


            If Not oWord Is Nothing Then
                If System.IO.File.Exists(Doc_Path & strDOC_NAME) Then
                    With oWord

                        DestDoc = .Documents.Open(Doc_Path & strDOC_NAME)
                        DestDoc = oWord.ActiveDocument
                        For Each dr In dtTable.Rows
                            FillUnencumberanceMemoTemplate(DestDoc, bolBreak, dr("FacilityID"), strDocType, strTempPath, pOwner, dr("TecEventID"), dr("FinEventID"), dr("CommitmentID"), dr("CommitAdjustID"))
                            bolBreak = True
                        Next
                        DestDoc.Save()
                        .Visible = True
                    End With
                End If

                Try
                    'If nFinEventID > 0 Then
                    '    nOwningEntity = nFinEventID
                    '    nEntityID = 32
                    'End If
                    'If nCommitmentID > 0 Then
                    '    nOwningEntity = nCommitmentID
                    '    nEntityID = 33
                    'End If
                    'If nReimbursementID > 0 Then
                    '    nOwningEntity = nReimbursementID
                    '    nEntityID = 35
                    'End If
                    UIUtilsGen.SaveDocument(0, 0, strDOC_NAME, strDocType, Doc_Path, strDocDesc, nModuleID, 0, 0, 0)
                Catch ex As Exception
                    Throw ex
                End Try
            End If
            oWord = Nothing

        Catch ex As Exception
            Throw ex
        Finally
            SetCursorType(System.Windows.Forms.Cursors.Default)
        End Try
    End Function


    Friend Function GenerateFinancialPoRequest(ByVal dtTable As DataTable, ByVal strDocType As String, ByVal strDocName As String, ByVal strDocDesc As String, ByVal strTemplateName As String) As Boolean
        Dim strDOC_NAME As String = String.Empty
        Dim DestDoc As Word.Document
        Dim strToday As String = String.Empty
        Dim bolBreak As Boolean = False
        Dim nEntityID As Int16
        Dim nOwningEntity As Int64
        Dim dr As DataRow
        nModuleID = UIUtilsGen.ModuleID.Financial
        Try
            If DOC_PATH = "\" Then
                Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
            End If


            strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm")) + CStr(Format(Now, "ss"))
            strDOC_NAME = "FIN_" + strDocName.Trim.ToString + "_" + strToday + ".doc"
            Dim strTempPath As String = TmpltPath & "Financial\" & strTemplateName
            If Not System.IO.File.Exists(strTempPath) Then
                Throw New Exception("File Not Found: " + strTempPath)
            End If
            System.IO.File.Copy(strTempPath, Doc_Path & strDOC_NAME)

            Dim oWord As Word.Application = MusterContainer.GetWordApp



            If Not oWord Is Nothing Then
                If System.IO.File.Exists(Doc_Path & strDOC_NAME) Then
                    With oWord

                        DestDoc = .Documents.Open(Doc_Path & strDOC_NAME)
                        DestDoc = oWord.ActiveDocument
                        For Each dr In dtTable.Select("", "ID")


                            If dr("REIMBURSE ERAC").ToString.ToUpper = "FALSE" Then
                                dr("ERACNAME") = String.Empty
                                dr("ERACNUM") = String.Empty
                            End If
                            FillPORequestMemoTemplate(DestDoc, bolBreak, dr("Fac_ID"), strDocType, strTempPath, dr("FacName").ToString, IIf(dr("ERACNAME").ToString.Length > 1, dr("ERACNAME").ToString, dr("Vendor_Name").ToString), IIf(dr("ERACNUM").ToString.Length > 1, dr("ERACNUM").ToString, dr("Vendor_Number").ToString), String.Format("{0:C}", dr("Balance")), dr("OldPO").ToString)
                            bolBreak = True
                        Next
                        DestDoc.Save()
                        .Visible = True
                    End With
                End If
                Try
                    UIUtilsGen.SaveDocument(0, 0, strDOC_NAME, strDocType, Doc_Path, strDocDesc, nModuleID, 0, 0, 0)
                Catch ex As Exception
                    Throw ex
                End Try
            End If
            oWord = Nothing

        Catch ex As Exception
            Throw ex
        Finally
            SetCursorType(System.Windows.Forms.Cursors.Default)
        End Try
    End Function

    Private Function FillUnencumberanceMemoTemplate(ByRef doc As Word.Document, ByVal bolBreak As Boolean, ByVal FacilityID As Integer, ByVal strDocType As String, ByVal strTemplateName As String, Optional ByRef pOwner As MUSTER.BusinessLogic.pOwner = Nothing, Optional ByVal nTechEventID As Int64 = 0, Optional ByVal nFinEventID As Int64 = 0, Optional ByVal nCommitmentID As Integer = 0, Optional ByVal nAdjustmentID As Integer = 0, Optional ByVal nOwnerID As Integer = 0) As Boolean
        Dim strContactName As String
        Dim bolNeedCapDocs As Boolean = False


        Dim nFacId As Integer = 0
        Dim sCommitmentGrandTotal As Double
        Dim oFacInfo As MUSTER.Info.FacilityInfo
        Dim oOwnerInfo As MUSTER.Info.OwnerInfo
        Dim oPersonaInfo As MUSTER.Info.PersonaInfo
        Dim oAddressInfo As MUSTER.Info.AddressInfo
        Dim oUserPM As New MUSTER.BusinessLogic.pUser
        Dim oLustEvent As New MUSTER.BusinessLogic.pLustEvent
        Dim oFinancialEvent As New MUSTER.BusinessLogic.pFinancial
        Dim oERACContact As New MUSTER.BusinessLogic.pCompany
        Dim oIRACContact As New MUSTER.BusinessLogic.pCompany
        Dim oContact As New MUSTER.BusinessLogic.pContactDatum
        Dim oVendor As New MUSTER.BusinessLogic.pContactDatum
        Dim oContactInfo As New MUSTER.Info.ContactDatumInfo
        Dim oCommitment As New MUSTER.BusinessLogic.pFinancialCommitment
        Dim oAdjustment As New MUSTER.BusinessLogic.pFinancialCommitAdjustment
        Dim oReimbursement As New MUSTER.BusinessLogic.pFinancialReimbursement
        Dim oInvoice As New MUSTER.BusinessLogic.pFinancialInvoice
        Dim colInvoice As New MUSTER.Info.FinancialInvoiceCollection
        Dim oInvoiceInfo As New MUSTER.Info.FinancialInvoiceInfo
        Dim oActivity As New MUSTER.BusinessLogic.pFinancialActivity
        Dim strInvoiceNumbers As String
        Dim strDeductionReasons As String
        Dim nPaidAmount As Double
        Dim nRequestAmount As Double
        Dim xArray As Array
        Dim i As Int16

        Dim xCostFormat As New CostFormat
        Dim oProperty As New MUSTER.BusinessLogic.pProperty
        Dim strVendorName As String
        Dim strVendorID As String = String.Empty
        Dim strPayeeVendorName As String = String.Empty
        Dim strVendorAddress1 As String = String.Empty
        Dim strVendorAddress2 As String = String.Empty
        Dim strVendorCity As String = String.Empty
        Dim strVendorState As String = String.Empty
        Dim strVendorZip As String = String.Empty
        Dim bolIsPerson As Boolean = False
        Dim nContactID As Integer
        Dim colParams As New Specialized.NameValueCollection
        Dim strKey As String = String.Empty
        Dim strValue As String = String.Empty
        Dim pcontactstruct As New BusinessLogic.pContactStruct
        Dim dtEracContact As DataSet
        Dim v As DataView

        dtEracContact = pcontactstruct.GetFilteredContacts(nFinEventID, nModuleID, , True)

        pcontactstruct = Nothing

        Try

            If Not dtEracContact Is Nothing AndAlso dtEracContact.Tables.Count > 0 Then
                dtEracContact.Tables(0).DefaultView.RowFilter = "ContactType = 42"
                dtEracContact.Tables(0).DefaultView.Sort = "IsPerson Desc"
                v = dtEracContact.Tables(0).DefaultView
            Else
                v = New DataView
            End If


            With doc
                doc.Activate()
                If bolBreak Then
                    MusterContainer.WordApp.Selection.EndKey(Unit:=Word.WdUnits.wdStory)
                    doc.Application.Selection.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage)
                    ' insert file
                    doc.Application.Selection.InsertFile(FILENAME:=strTemplateName, ConfirmConversions:=False, Link:=False, Attachment:=False)
                End If

                strContactName = ""
                oOwnerInfo = pOwner.Retrieve(nOwnerID, "SELF")
                oFacInfo = pOwner.Facilities.Retrieve(pOwner.OwnerInfo, FacilityID, "SELF", "FACILITY")
                If Not bolBreak Then
                    colParams.Add("<Title>", strDocType.ToString)
                End If
                colParams.Add("<DATE>", Format(Now, "MMMM d, yyyy"))


                oLustEvent.Retrieve(nTechEventID)
                oUserPM.Retrieve(oLustEvent.PM)

                strVendorName = ""
                'Dim strContactName As String = ""
                If nFinEventID > 0 Then
                    Dim drRow As DataRow
                    Dim bolVendorPayee As Boolean = False

                    oFinancialEvent.Retrieve(nFinEventID)

                    If nCommitmentID > 0 Then
                        oCommitment.Retrieve(nCommitmentID)
                        colParams.Add("<PONumber>", oCommitment.PONumber.ToString)
                    End If

                    Dim ERACPayee As Boolean = oCommitment.ReimburseERAC

                    If oLustEvent.ERAC <> 0 Then
                        Dim dtErac As DataTable

                        dtErac = oFinancialEvent.GetProjectEngineer(oLustEvent.EVENTSEQUENCE, FacilityID)


                        oERACContact.Retrieve(oLustEvent.ERAC)

                        If (Not dtErac Is Nothing AndAlso (dtErac.Rows.Count > 0 OrElse Not oERACContact Is Nothing)) OrElse (Not v Is Nothing AndAlso dtEracContact.Tables.Count > 0 AndAlso v.Count > 0) Then

                            If v.Count > 0 Then

                                strPayeeVendorName = IIf(v.Item(0)("Assoc Company") Is DBNull.Value OrElse v.Item(0)("Assoc Company") = "", v.Item(0)("Contact_name"), v.Item(0)("Assoc Company"))

                                colParams.Add("<ERAC>", IIf(v.Item(0)("Assoc Company") Is DBNull.Value OrElse v.Item(0)("Assoc Company") = "", v.Item(0)("Contact_name"), v.Item(0)("Assoc Company")))

                                If (v.Item(0)("Assoc Company") Is DBNull.Value OrElse v.Item(0)("Assoc Company") = "") Then

                                    colParams.Add("<ERAC Contact>, <ERAC>", "Cut")
                                    colParams.Add("<ERAC Company>", v.Item(0)("Contact_Name"))
                                    colParams.Add("<ERAC Contact>", "")


                                Else

                                    colParams.Add("<ERAC Company>", v.Item(0)("Assoc Company"))
                                    colParams.Add("<ERAC Contact>", v.Item(0)("Contact_Name"))


                                End If

                            ElseIf oERACContact.ID > 0 Then
                                colParams.Add("<ERAC Company>", oERACContact.COMPANY_NAME)
                                colParams.Add("<ERAC>", oERACContact.COMPANY_NAME)
                                strPayeeVendorName = oERACContact.COMPANY_NAME

                            End If

                            If v.Count = 0 AndAlso dtErac.Rows.Count > 0 Then

                                If (dtErac.Rows(0).Item("PRO_ENGIN") = "") Then
                                    If (dtErac.Rows(0).Item("PRO_GEOLO") = "") Then
                                        colParams.Add("<ERAC Contact>", dtErac.Rows(0).Item("PRO_GEOLO"))
                                    Else
                                        colParams.Add("<ERAC Contact>", "")
                                    End If

                                Else
                                    colParams.Add("<ERAC Contact>", dtErac.Rows(0).Item("PRO_ENGIN"))
                                End If
                            End If




                            If (v.Count = 0 AndAlso dtErac.Rows.Count > 0) OrElse (dtErac.Rows.Count > 0 AndAlso v.Count > 0 AndAlso ((v.Item(0)("ADDRESS_ONE") Is DBNull.Value OrElse v.Item(0)("ADDRESS_ONE") = "") _
                                                   And (v.Item(0)("ADDRESS_TWO") Is DBNull.Value OrElse v.Item(0)("ADDRESS_TWO") = "") _
                                                   And (v.Item(0)("City") Is DBNull.Value OrElse v.Item(0)("City") = "") _
                                                   And (v.Item(0)("STATE") Is DBNull.Value OrElse v.Item(0)("STATE") = ""))) Then

                                colParams.Add("<ERAC Address1>", dtErac.Rows(0).Item("ADDRESS_LINE_ONE"))
                                If ERACPayee Then
                                    strVendorName = strPayeeVendorName
                                    strVendorCity = dtErac.Rows(0).Item("CITY")
                                    strVendorAddress1 = dtErac.Rows(0).Item("ADDRESS_LINE_ONE")
                                    If dtErac.Rows(0).Item("ADDRESS_LINE_TWO") Is DBNull.Value Then
                                        strVendorAddress2 = ""
                                    Else
                                        strVendorAddress2 = dtErac.Rows(0).Item("ADDRESS_LINE_TWO")
                                    End If

                                    strVendorState = dtErac.Rows(0).Item("STATE")
                                    strVendorZip = dtErac.Rows(0).Item("ZIP")

                                End If
                                If Not dtErac.Rows(0).Item("ADDRESS_LINE_TWO") Is System.DBNull.Value Then
                                    If dtErac.Rows(0).Item("ADDRESS_LINE_TWO") Is DBNull.Value Then
                                        colParams.Add("<ERAC Address2>", dtErac.Rows(0).Item("CITY") + ", " + _
                                                                            dtErac.Rows(0).Item("STATE") + " " + _
                                                                            dtErac.Rows(0).Item("ZIP"))
                                        colParams.Add("<ERAC City/State/Zip>", "")
                                    Else
                                        colParams.Add("<ERAC Address2>", dtErac.Rows(0).Item("ADDRESS_LINE_TWO"))
                                        colParams.Add("<ERAC City/State/Zip>", dtErac.Rows(0).Item("CITY") + ", " + _
                                                                                dtErac.Rows(0).Item("STATE") + " " + _
                                                                                dtErac.Rows(0).Item("ZIP"))
                                    End If
                                Else
                                    colParams.Add("<ERAC Address2>", dtErac.Rows(0).Item("CITY") + ", " + _
                                                                        dtErac.Rows(0).Item("STATE") + " " + _
                                                                        dtErac.Rows(0).Item("ZIP"))
                                    colParams.Add("<ERAC City/State/Zip>", "")
                                End If
                            ElseIf v.Count > 0 Then

                                If ERACPayee Then
                                    strVendorName = strPayeeVendorName
                                    strVendorCity = v.Item(0)("CITY")
                                    strVendorAddress1 = v.Item(0)("ADDRESS_ONE")
                                    strVendorAddress2 = v.Item(0)("ADDRESS_TWO")
                                    strVendorState = v.Item(0)("STATE")
                                    strVendorZip = v.Item(0)("ZIP")

                                End If


                                colParams.Add("<ERAC Address1>", v.Item(0)("ADDRESS_ONE"))

                                If v.Item(0)("ADDRESS_TWO") Is DBNull.Value OrElse v.Item(0)("ADDRESS_TWO") Is DBNull.Value Then
                                    colParams.Add("<ERAC Address2>", v.Item(0)("CITY") + " " + _
                                                                        v.Item(0)("STATE") + " " + _
                                                                        v.Item(0)("ZIP"))
                                    colParams.Add("<ERAC City/State/Zip>", "")
                                Else
                                    colParams.Add("<ERAC Address2>", v.Item(0)("ADDRESS_TWO"))
                                    colParams.Add("<ERAC City/State/Zip>", v.Item(0)("CITY") + ", " + _
                                                                            v.Item(0)("STATE") + " " + _
                                                                            v.Item(0)("ZIP"))
                                End If
                            Else
                                colParams.Add("<ERAC Contact>", String.Empty)
                                colParams.Add("<ERAC Address1>", String.Empty)
                                colParams.Add("<ERAC Address2>", String.Empty)
                                colParams.Add("<ERAC City/State/Zip>", String.Empty)
                            End If

                        Else
                            'colParams.Add("<ERAC>", String.Empty)
                            colParams.Add("<ERAC Contact>", String.Empty)
                            colParams.Add("<ERAC Address1>", String.Empty)
                            colParams.Add("<ERAC Address2>", String.Empty)
                            colParams.Add("<ERAC City/State/Zip>", String.Empty)
                        End If


                    Else
                        colParams.Add("<ERAC>", String.Empty)
                        colParams.Add("<ERAC Contact>", String.Empty)
                        colParams.Add("<ERAC Address1>", String.Empty)
                        colParams.Add("<ERAC Address2>", String.Empty)
                        colParams.Add("<ERAC City/State/Zip>", String.Empty)

                    End If
                    If oLustEvent.IRAC > 0 Then
                        oIRACContact.Retrieve(oLustEvent.IRAC)
                        colParams.Add("<IRAC>", oIRACContact.COMPANY_NAME)
                    Else
                        colParams.Add("<IRAC>", String.Empty)
                    End If

                    strVendorName = ""
                    'Dim strContactName As String = ""

                    oFinancialEvent.Retrieve(nFinEventID)
                    Dim dtVendor As DataTable = GetXHAndXLContacts(nFinEventID, UIUtilsGen.EntityTypes.FinancialEvent, UIUtilsGen.ModuleID.Financial, , 1) 'oVendor.GetByID(oFinancialEvent.VendorID, nFinEventID, 616)
                    If dtVendor.Rows.Count > 0 Then
                        For Each drRow In dtVendor.Rows
                            If drRow.Item("Type") = EnumContactType.XH Then
                                If (Not drRow.Item("AssocCompany") Is System.DBNull.Value) AndAlso (drRow.Item("AssocCompany").ToString.Trim.Length <> 0) Then
                                    If drRow.Item("IsPerson") = True Then
                                        strContactName = drRow.Item("CONTACT_Name")
                                        'colParams.Add("<ContactName>", drRow.Item("CONTACT_Name")) 'IIf(drRow.Item("Title") = String.Empty, "", drRow.Item("Title") + " ") + drRow.Item("First_Name") + " " + IIf(drRow.Item("Middle_Name") = String.Empty, "", drRow.Item("Middle_Name") + " ") + drRow.Item("Last_Name") + IIf(drRow.Item("Suffix") = String.Empty, "", " " + drRow.Item("Suffix")))
                                        'colParams.Add("<Salutation>", "Dear " & drRow.Item("CONTACT_Name")) 'IIf(drRow.Item("Title") = String.Empty, "", drRow.Item("Title") + " ") + drRow.Item("First_Name") + " " + IIf(drRow.Item("Middle_Name") = String.Empty, "", drRow.Item("Middle_Name") + " ") + drRow.Item("Last_Name") + IIf(drRow.Item("Suffix") = String.Empty, "", " " + drRow.Item("Suffix")))
                                    End If
                                    strVendorName = drRow.Item("AssocCompany")
                                    strVendorID = IIf(drRow.Item("VENDOR_NUMBER") Is DBNull.Value, String.Empty, drRow.Item("VENDOR_NUMBER"))
                                Else
                                    'If drRow.Item("IsPerson") = True Then
                                    strVendorName = drRow.Item("CONTACT_Name") 'IIf(drRow.Item("Title") = String.Empty, "", drRow.Item("Title") + " ") + drRow.Item("First_Name") + " " + IIf(drRow.Item("Middle_Name") = String.Empty, "", drRow.Item("Middle_Name") + " ") + drRow.Item("Last_Name") + IIf(drRow.Item("Suffix") = String.Empty, "", " " + drRow.Item("Suffix"))
                                    strVendorID = IIf(drRow.Item("VENDOR_NUMBER") Is DBNull.Value, String.Empty, drRow.Item("VENDOR_NUMBER"))
                                    'Else
                                    '    strVendorName = drRow.Item("Company_Name")
                                    'End If
                                    'colParams.Add("<ContactName>", String.Empty)
                                    'colParams.Add("<Salutation>", "Dear " & strVendorName)
                                End If
                                If drRow.Item("IsPerson") = True Then
                                    bolIsPerson = True
                                    nContactID = Integer.Parse(drRow.Item("ContactID"))
                                End If

                                If ERACPayee Then

                                    colParams.Add("<VendorERAC>", strPayeeVendorName)

                                Else
                                    colParams.Add("<VendorERAC>", String.Empty)
                                End If

                                colParams.Add("<Vendor>", strVendorName)

                                If colParams.Item("<VendorAddress1>") Is Nothing Then

                                    colParams.Add("<VendorAddress1>", IIf(drRow.Item("Address_One") Is DBNull.Value, "", drRow.Item("Address_One")))

                                    If drRow.Item("Address_Two") Is DBNull.Value OrElse drRow.Item("Address_Two") = String.Empty Then
                                        colParams.Add("<VendorAddress2>", drRow.Item("City") & ", " & _
                                                                            drRow.Item("State") & " " & _
                                                                            drRow.Item("Zip"))
                                        colParams.Add("<VendorCityStateZip>", "")
                                    Else
                                        colParams.Add("<VendorAddress2>", drRow.Item("Address_Two"))
                                        colParams.Add("<VendorCityStateZip>", drRow.Item("City") & ", " & _
                                                                                drRow.Item("State") & " " & _
                                                                                drRow.Item("Zip"))
                                    End If
                                End If

                                If Not bolVendorPayee And Not ERACPayee Then
                                    strPayeeVendorName = strVendorName
                                    strVendorAddress1 = IIf(drRow.Item("Address_One") Is DBNull.Value, "", drRow.Item("Address_One"))
                                    strVendorAddress2 = IIf(drRow.Item("Address_Two") Is DBNull.Value, "", drRow.Item("Address_Two"))
                                    strVendorCity = IIf(drRow.Item("City") Is DBNull.Value, "", drRow.Item("City"))
                                    strVendorState = IIf(drRow.Item("State") Is DBNull.Value, "", drRow.Item("State"))
                                    strVendorZip = IIf(drRow.Item("Zip") Is DBNull.Value, "", drRow.Item("Zip"))
                                End If
                            ElseIf UCase(drRow.Item("ContactType")) = UCase("Financial Payee") AndAlso Not ERACPayee Then
                                bolVendorPayee = True
                                strPayeeVendorName = IIf(drRow.Item("AssocCompany") Is DBNull.Value OrElse drRow.Item("AssocCompany") = String.Empty, drRow.Item("CONTACT_Name"), drRow.Item("AssocCompany"))
                                strVendorID = IIf(drRow.Item("VENDOR_NUMBER") Is DBNull.Value, String.Empty, drRow.Item("VENDOR_NUMBER"))
                                strVendorAddress1 = IIf(drRow.Item("Address_One") Is DBNull.Value, "", drRow.Item("Address_One"))
                                strVendorAddress2 = IIf(drRow.Item("Address_Two") Is DBNull.Value, "", drRow.Item("Address_Two"))
                                strVendorCity = IIf(drRow.Item("City") Is DBNull.Value, "", drRow.Item("City"))
                                strVendorState = IIf(drRow.Item("State") Is DBNull.Value, "", drRow.Item("State"))
                                strVendorZip = IIf(drRow.Item("Zip") Is DBNull.Value, "", drRow.Item("Zip"))
                            ElseIf UCase(drRow.Item("ContactType")) = UCase("Financial Representative") Then
                            End If

                            If drRow.Item("Type") = EnumContactType.XL Then
                                If drRow.Item("IsPerson") = True Then
                                    bolIsPerson = True
                                    nContactID = Integer.Parse(drRow.Item("ContactID"))
                                    strContactName = drRow.Item("CONTACT_Name")
                                End If
                            End If

                        Next
                    End If


                    'Sets the ERAC Payee
                    If ERACPayee Then


                        oContact.Retrieve(nContactID)

                        If v.Count > 0 AndAlso Not (v.Item(0)("Vendor_Number") Is DBNull.Value OrElse v.Item(0)("Vendor_Number") = "" OrElse v.Item(0)("Vendor_Number") = "0") Then
                            strVendorID = v.Item(0)("Vendor_Number")

                        ElseIf oContact Is Nothing OrElse nContactID = 0 Then
                            strVendorID = "<Please Associate an ERAC contact with a vendor # and regenerate this notice>"
                        ElseIf oContact.VendorNumber.Length > 0 Then
                            strVendorID = oContact.VendorNumber
                        Else
                            strVendorID = "<Please Associate an ERAC contact with a vendor # and regenerate this notice>"
                        End If
                    ElseIf strVendorID Is Nothing OrElse strVendorID.Length = 0 Then
                        strVendorID = "<Please Associate an ERAC contact or financial vendor with a vendor #  and regenerate this notice>"
                    End If

                    colParams.Add("<PayeeVendor>", strPayeeVendorName)
                    colParams.Add("<VendorNumber>", strVendorID)
                    colParams.Add("<PayeeVendorAddress1>", strVendorAddress1)
                    If strVendorAddress2 = String.Empty Then
                        colParams.Add("<PayeeVendorAddress2>", strVendorCity & IIf(strVendorCity.Length > 0, ", ", "") & _
                                                                strVendorState & " " & _
                                                                strVendorZip)
                        colParams.Add("<PayeeVendorCityStateZip>", "")
                    Else
                        colParams.Add("<PayeeVendorAddress2>", strVendorAddress2)
                        colParams.Add("<PayeeVendorCityStateZip>", strVendorCity & IIf(strVendorCity.Length > 0, ", ", "") & _
                                                                    strVendorState & " " & _
                                                                    strVendorZip)
                    End If

                Else
                    If colParams.Item("<VendorAddress1>") Is Nothing Then
                        colParams.Add("<Vendor>", String.Empty)
                        colParams.Add("<VendorAddress1>", String.Empty)
                        colParams.Add("<VendorAddress2>", String.Empty)
                        colParams.Add("<VendorCityStateZip>", String.Empty)
                        colParams.Add("<Contact Name>", String.Empty)
                    End If


                End If

                Dim dsSalutation As DataSet
                If bolIsPerson Then

                    dsSalutation = oContact.GetContactLastName(nContactID)

                End If
                If strContactName <> String.Empty Then
                    colParams.Add("<ContactName>", strContactName)
                    If bolIsPerson Then
                        colParams.Add("<Salutation>", strContactName)
                    Else
                        colParams.Add("<Salutation>", strContactName)
                    End If

                ElseIf strVendorName <> String.Empty Then
                    colParams.Add("<ContactName>", String.Empty)
                    If bolIsPerson Then
                        colParams.Add("<Salutation>", strVendorName)
                    Else
                        colParams.Add("<Salutation>", strVendorName)
                    End If
                Else
                    If oOwnerInfo.OrganizationID > 0 Then
                        oPersonaInfo = pOwner.Organization
                        'colParams.Add("<VendorName>", pOwner.BPersona.Company)
                        colParams.Add("<Vendor>", pOwner.BPersona.Company)
                        colParams.Add("<Salutation>", pOwner.BPersona.Company.Trim)
                    Else
                        oPersonaInfo = pOwner.Persona
                        colParams.Add("<Vendor>", pOwner.BPersona.FirstName.Trim & IIf(pOwner.BPersona.MiddleName.Trim.Length > 0, " ", "") & pOwner.BPersona.MiddleName.Trim & IIf(pOwner.BPersona.LastName.Trim.Length > 0, " ", "") & pOwner.BPersona.LastName.Trim & IIf(pOwner.BPersona.Suffix.Trim.Length > 0, " ", "") & pOwner.BPersona.Suffix.Trim)
                        colParams.Add("<Salutation>", pOwner.BPersona.FirstName.Trim & IIf(pOwner.BPersona.MiddleName.Trim.Length > 0, " ", "") & pOwner.BPersona.MiddleName.Trim & IIf(pOwner.BPersona.LastName.Trim.Length > 0, " ", "") & pOwner.BPersona.LastName.Trim & IIf(pOwner.BPersona.Suffix.Trim.Length > 0, " ", "") & pOwner.BPersona.Suffix.Trim)
                    End If
                    colParams.Add("<ContactName>", String.Empty)
                    'Else
                    '    colParams.Add("<Vendor>", strVendorName)
                End If

                If oOwnerInfo.OrganizationID > 0 Then
                    oPersonaInfo = pOwner.Organization
                    'colParams.Add("<Salutation>", pOwner.BPersona.Company.Trim)
                    colParams.Add("<OwnerName>", pOwner.BPersona.Company)
                Else
                    oPersonaInfo = pOwner.Persona
                    'colParams.Add("<Salutation>", "Dear " & pOwner.BPersona.Title.Trim & IIf(pOwner.BPersona.Title.Trim.Length > 0, " ", "") & pOwner.BPersona.FirstName.Trim & IIf(pOwner.BPersona.MiddleName.Trim.Length > 0, " ", "") & pOwner.BPersona.MiddleName.Trim & IIf(pOwner.BPersona.LastName.Trim.Length > 0, " ", "") & pOwner.BPersona.LastName.Trim & IIf(pOwner.BPersona.Suffix.Trim.Length > 0, " ", "") & pOwner.BPersona.Suffix.Trim & "")
                    colParams.Add("<OwnerName>", pOwner.BPersona.FirstName.Trim & IIf(pOwner.BPersona.MiddleName.Trim.Length > 0, " ", "") & pOwner.BPersona.MiddleName.Trim & IIf(pOwner.BPersona.LastName.Trim.Length > 0, " ", "") & pOwner.BPersona.LastName.Trim & IIf(pOwner.BPersona.Suffix.Trim.Length > 0, " ", "") & pOwner.BPersona.Suffix.Trim)
                End If

                If strVendorName = "" AndAlso colParams.Item("<VendorAddress1>") Is Nothing Then
                    oAddressInfo = pOwner.Address()
                    'colParams.Add("<Company Address>", oAddressInfo.AddressLine1)
                    colParams.Add("<VendorAddress1>", oAddressInfo.AddressLine1)

                    If oAddressInfo.AddressLine2 = String.Empty Then
                        colParams.Add("<VendorAddress2>", oAddressInfo.City & ", " & oAddressInfo.State.TrimEnd & " " & oAddressInfo.Zip)
                        colParams.Add("<VendorCityStateZip>", "")
                    Else
                        colParams.Add("<VendorAddress2>", oAddressInfo.AddressLine2)
                        colParams.Add("<VendorCityStateZip>", oAddressInfo.City & ", " & oAddressInfo.State.TrimEnd & " " & oAddressInfo.Zip)
                    End If
                    'colParams.Add("<OwnerCityStateZip>", oAddressInfo.City & ", " & oAddressInfo.State.TrimEnd & " " & oAddressInfo.Zip)
                End If
                oAddressInfo = pOwner.Addresses.Retrieve(oFacInfo.AddressID)
                colParams.Add("<FacilityName>", oFacInfo.Name)
                'colParams.Add("<FacilityStreetAddress>", oAddressInfo.AddressLine1 & " " & IIf(oAddressInfo.AddressLine2 = String.Empty, "", "; " & oAddressInfo.AddressLine2))
                colParams.Add("<Facility Address1>", oAddressInfo.AddressLine1)
                If oAddressInfo.AddressLine2 = String.Empty Then
                    colParams.Add("<Facility Address2>", oAddressInfo.City & ", " & oAddressInfo.State.TrimEnd & " " & oAddressInfo.Zip)
                    colParams.Add("<Facility City/State/Zip>", "")
                Else
                    colParams.Add("<Facility Address2>", oAddressInfo.AddressLine2)
                    colParams.Add("<Facility City/State/Zip>", oAddressInfo.City & ", " & oAddressInfo.State.TrimEnd & " " & oAddressInfo.Zip)
                End If
                'colParams.Add("<FacilityCityStateZip>", oAddressInfo.City.TrimEnd & ", " & oAddressInfo.State.TrimEnd & " " & oAddressInfo.Zip.TrimEnd)
                colParams.Add("<FacilityID>", oFacInfo.ID.ToString)



                colParams.Add("<ProjectManager>", oUserPM.Name)
                colParams.Add("<ProjectManagerPhone>", oUserPM.PhoneNumber)


                If nCommitmentID > 0 Then
                    oCommitment.Retrieve(nCommitmentID)
                    colParams.Add("<DueDateText>", oCommitment.DueDateStatement)
                    colParams.Add("<CommitmentAmount>", FormatNumber(sCommitmentGrandTotal, 2, TriState.True, TriState.UseDefault, TriState.True))
                    colParams.Add("<PurchaseOrderNumber>", oCommitment.PONumber)
                    colParams.Add("<SOW DATE>", Format(oCommitment.SOWDate, "MMMM d, yyyy"))
                Else
                    colParams.Add("<DueDateText>", String.Empty)
                    colParams.Add("<CommitmentAmount>", String.Empty)
                    colParams.Add("<PurchaseOrderNumber>", String.Empty)
                End If

                If nAdjustmentID > 0 Then
                    oAdjustment.Retrieve(nAdjustmentID)
                    colParams.Add("<ChangeOrderAmount>", FormatNumber(oAdjustment.AdjustAmount, 2, TriState.True, TriState.False, TriState.True))
                Else
                    colParams.Add("<ChangeOrderAmount>", String.Empty)
                End If


                ' Find and Replace the TAGs with Values.
                For j As Integer = 0 To colParams.Count - 1
                    strKey = colParams.Keys(j).ToString
                    strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                    .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)
                Next

            End With
        Catch ex As Exception
            Throw ex
        End Try
    End Function





    Private Function FillPORequestMemoTemplate(ByRef doc As Word.Document, ByVal bolBreak As Boolean, ByVal FacilityID As String, ByVal strDocType As String, ByVal strTemplateName As String, ByVal strFacName As String, ByVal strVendorName As String, ByVal strVendorNum As String, ByVal dblBalance As String, ByVal oldPo As String) As Boolean
        Dim strContactName As String
        Dim bolNeedCapDocs As Boolean = False


        Dim nFacId As Integer = 0
        Dim nPaidAmount As Double
        Dim nRequestAmount As Double

        Dim strKey As String
        Dim strValue As String


        Dim xArray As Array
        Dim i As Int16

        Dim oProperty As New MUSTER.BusinessLogic.pProperty

        Dim colParams As New Specialized.NameValueCollection


        Try
            With doc
                doc.Activate()
                If bolBreak Then
                    MusterContainer.WordApp.Selection.EndKey(Unit:=Word.WdUnits.wdStory)
                    doc.Application.Selection.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage)
                    ' insert file
                    doc.Application.Selection.InsertFile(FILENAME:=strTemplateName, ConfirmConversions:=False, Link:=False, Attachment:=False)
                End If


                colParams.Add("<PayeeVendor>", strVendorName)
                colParams.Add("<Approved Date>", Date.Now.ToString("MMM d, yyyy"))

                colParams.Add("<VendorNumber>", strVendorNum)
                colParams.Add("<FacilityID>", FacilityID)
                colParams.Add("<FacilityName>", strFacName)
                colParams.Add("<PONumber>", String.Format("Old PO Number: {0}", oldPo))
                colParams.Add("<CommitmentAmount>", dblBalance)

                ' Find and Replace the TAGs with Values.
                For j As Integer = 0 To colParams.Count - 1
                    strKey = colParams.Keys(j).ToString
                    strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                    .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)
                Next

            End With
        Catch ex As Exception
            Throw ex
        End Try
    End Function











    Friend Function GenerateFinancialLetter(ByVal FacilityID As Integer, ByVal strDocType As String, ByVal strDocName As String, ByVal strDocDesc As String, ByVal strTemplateName As String, Optional ByRef pOwner As MUSTER.BusinessLogic.pOwner = Nothing, Optional ByVal nTechEventID As Int64 = 0, Optional ByVal nFinEventID As Int64 = 0, Optional ByVal nCommitmentID As Integer = 0, Optional ByVal nReimbursementID As Integer = 0, Optional ByVal nAdjustmentID As Integer = 0, Optional ByVal nInvoiceID As Integer = 0, Optional ByVal nOwnerID As Integer = 0, Optional ByVal eventID As Int64 = 0, Optional ByVal eventSequence As Integer = 0, Optional ByVal eventType As Integer = 0, Optional ByVal contactID As Integer = 0) As Boolean
        Dim strContactName As String
        Dim strContactFirstName As String
        Dim bolNeedCapDocs As Boolean = False
        Dim strDOC_NAME As String = String.Empty
        Dim strToday As String = String.Empty
        Dim nFacId As Integer = 0
        Dim sCommitmentGrandTotal As Double
        Dim oFacInfo As MUSTER.Info.FacilityInfo
        Dim oOwnerInfo As MUSTER.Info.OwnerInfo
        Dim oPersonaInfo As MUSTER.Info.PersonaInfo
        Dim oAddressInfo As MUSTER.Info.AddressInfo
        Dim pAddressInfo As MUSTER.Info.AddressInfo
        Dim nEntityID As Int16
        Dim nOwningEntity As Int64

        Dim oUserPM As New MUSTER.BusinessLogic.pUser
        Dim oLustEvent As New MUSTER.BusinessLogic.pLustEvent
        Dim oFinancialEvent As New MUSTER.BusinessLogic.pFinancial  'MUSTER.BusinessLogic.pFinancial
        Dim oERACContact As New MUSTER.BusinessLogic.pCompany
        Dim oIRACContact As New MUSTER.BusinessLogic.pCompany
        Dim oContact As New MUSTER.BusinessLogic.pContactDatum
        'Dim oVendor As New MUSTER.BusinessLogic.pContactDatum
        Dim oContactInfo As New MUSTER.Info.ContactDatumInfo
        Dim oCommitment As New MUSTER.BusinessLogic.pFinancialCommitment
        Dim oAdjustment As New MUSTER.BusinessLogic.pFinancialCommitAdjustment
        Dim oReimbursement As New MUSTER.BusinessLogic.pFinancialReimbursement
        Dim oInvoice As New MUSTER.BusinessLogic.pFinancialInvoice
        Dim colInvoice As New MUSTER.Info.FinancialInvoiceCollection
        Dim oInvoiceInfo As New MUSTER.Info.FinancialInvoiceInfo
        Dim oActivity As New MUSTER.BusinessLogic.pFinancialActivity
        Dim oUserExecutiveDirectorInfo As New MUSTER.Info.UserInfo
        Dim strInvoiceNumbers As String
        Dim strInvoiceComments As String
        Dim strDeductionReasons As String
        Dim nPaidAmount As Double
        Dim nRequestAmount As Double
        Dim xArray As Array
        Dim i As Int16

        Dim xCostFormat As New CostFormat
        Dim oProperty As New MUSTER.BusinessLogic.pProperty
        Dim strFinRep As String = String.Empty
        Dim strVendorName As String
        Dim strVendorID As String = String.Empty
        Dim strPayeeVendorName As String = String.Empty
        Dim strVendorAddress1 As String = String.Empty
        Dim strVendorAddress2 As String = String.Empty
        Dim strVendorCity As String = String.Empty
        Dim strVendorState As String = String.Empty
        Dim strVendorZip As String = String.Empty
        Dim bolIsPerson As Boolean = False
        Dim nContactID As Integer
        Dim ERACPayee As Boolean = False

        Dim cc As String
        Dim ownerAddressID As Long


        Dim colParams As New Specialized.NameValueCollection
        nModuleID = UIUtilsGen.ModuleID.Financial
        Try


            Dim pcontactstruct As New BusinessLogic.pContactStruct
            Dim dtEracContact As DataSet
            Dim v As DataView

            dtEracContact = pcontactstruct.GetFilteredContacts(nFinEventID, nModuleID, , True)

            pcontactstruct = Nothing

            If Not dtEracContact Is Nothing AndAlso dtEracContact.Tables.Count > 0 Then
                dtEracContact.Tables(0).DefaultView.RowFilter = "ContactType = 42"
                dtEracContact.Tables(0).DefaultView.Sort = "IsPerson Desc"
                v = dtEracContact.Tables(0).DefaultView
            Else
                v = New DataView
            End If




            If DOC_PATH = "\" Then
                Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
            End If
            strContactName = ""

            strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm")) + CStr(Format(Now, "ss"))
            strDOC_NAME = "FIN_" + strDocName.Trim.ToString + "_" + CStr(Trim(FacilityID.ToString)) + "_" + strToday + ".doc"

            'Build NameValueCollection with Tags and Values.
            'oOwnerInfo = pOwner.Retrieve(pOwner.ID, "SELF")
            oOwnerInfo = pOwner.Retrieve(nOwnerID, "SELF")
            oFacInfo = pOwner.Facilities.Retrieve(pOwner.OwnerInfo, FacilityID, "SELF", "FACILITY")
            colParams.Add("<Title>", strDocType.ToString)
            colParams.Add("<DATE>", Format(Now, "MMMM d, yyyy"))


            oLustEvent.Retrieve(nTechEventID)
            oUserPM.Retrieve(oLustEvent.PM)


            If nReimbursementID > 0 Then
                oReimbursement.Retrieve(nReimbursementID)
            End If

            If nCommitmentID > 0 Then
                oCommitment.Retrieve(nCommitmentID)
                colParams.Add("<PONumber>", oCommitment.PONumber.ToString)
            End If

            ERACPayee = oCommitment.ReimburseERAC

            If oLustEvent.ERAC <> 0 OrElse ERACPayee Then
                Dim dtErac As DataTable

                dtErac = oFinancialEvent.GetProjectEngineer(oLustEvent.EVENTSEQUENCE, FacilityID)


                oERACContact.Retrieve(oLustEvent.ERAC)

                If (Not dtErac Is Nothing AndAlso (dtErac.Rows.Count > 0 OrElse Not oERACContact Is Nothing)) OrElse (Not v Is Nothing AndAlso dtEracContact.Tables.Count > 0 AndAlso v.Count > 0) Then

                    If v.Count > 0 Then

                        strPayeeVendorName = IIf(v.Item(0)("Assoc Company") Is DBNull.Value OrElse v.Item(0)("Assoc Company") = "", IIf(v.Item(0)("Contact_name") Is DBNull.Value, String.Empty, v.Item(0)("Contact_name")), v.Item(0)("Assoc Company"))
                        v.Item(0)("Contact_name") = IIf(v.Item(0)("Contact_name") Is DBNull.Value, String.Empty, v.Item(0)("Contact_name"))

                        colParams.Add("<ERAC>", IIf(v.Item(0)("Assoc Company") Is DBNull.Value OrElse v.Item(0)("Assoc Company") = "", IIf(v.Item(0)("Contact_name") Is DBNull.Value, String.Empty, v.Item(0)("Contact_name")), v.Item(0)("Assoc Company")))

                        If (v.Item(0)("Assoc Company") Is DBNull.Value OrElse v.Item(0)("Assoc Company") = "") Then

                            colParams.Add("<ERAC Contact>, <ERAC>", "Cut")
                            colParams.Add("<ERAC Company>", v.Item(0)("Contact_Name"))
                            colParams.Add("<ERAC Contact>", "")


                        Else

                            colParams.Add("<ERAC Company>", v.Item(0)("Assoc Company"))
                            colParams.Add("<ERAC Contact>", v.Item(0)("Contact_Name"))


                        End If

                    ElseIf oERACContact.ID > 0 Then
                        colParams.Add("<ERAC Company>", oERACContact.COMPANY_NAME)
                        colParams.Add("<ERAC>", oERACContact.COMPANY_NAME)
                        strPayeeVendorName = oERACContact.COMPANY_NAME

                    End If

                    If v.Count = 0 AndAlso dtErac.Rows.Count > 0 Then

                        If (dtErac.Rows(0).Item("PRO_ENGIN") = "") Then
                            If (dtErac.Rows(0).Item("PRO_GEOLO") = "") Then
                                colParams.Add("<ERAC Contact>", dtErac.Rows(0).Item("PRO_GEOLO"))
                            Else
                                colParams.Add("<ERAC Contact>", "")
                            End If

                        Else
                            colParams.Add("<ERAC Contact>", dtErac.Rows(0).Item("PRO_ENGIN"))
                        End If
                    End If




                    If (v.Count = 0 AndAlso dtErac.Rows.Count > 0) OrElse (dtErac.Rows.Count > 0 AndAlso v.Count > 0 AndAlso ((v.Item(0)("ADDRESS_ONE") Is DBNull.Value OrElse v.Item(0)("ADDRESS_ONE") = "") _
                                           And (v.Item(0)("ADDRESS_TWO") Is DBNull.Value OrElse v.Item(0)("ADDRESS_TWO") = "") _
                                           And (v.Item(0)("City") Is DBNull.Value OrElse v.Item(0)("City") = "") _
                                           And (v.Item(0)("STATE") Is DBNull.Value OrElse v.Item(0)("STATE") = ""))) Then

                        colParams.Add("<ERAC Address1>", dtErac.Rows(0).Item("ADDRESS_LINE_ONE"))
                        If ERACPayee Then
                            strVendorName = strPayeeVendorName
                            strVendorCity = IIf(dtErac.Rows(0).Item("CITY") Is DBNull.Value, String.Empty, dtErac.Rows(0).Item("CITY"))
                            strVendorAddress1 = dtErac.Rows(0).Item("ADDRESS_LINE_ONE")
                            If dtErac.Rows(0).Item("ADDRESS_LINE_TWO") Is DBNull.Value Then
                                strVendorAddress2 = ""
                            Else
                                strVendorAddress2 = dtErac.Rows(0).Item("ADDRESS_LINE_TWO")
                            End If

                            strVendorState = IIf(dtErac.Rows(0).Item("STATE") Is DBNull.Value, String.Empty, dtErac.Rows(0).Item("STATE"))
                            strVendorZip = IIf(dtErac.Rows(0).Item("ZIP") Is DBNull.Value, String.Empty, dtErac.Rows(0).Item("ZIP"))

                        End If
                        If Not dtErac.Rows(0).Item("ADDRESS_LINE_TWO") Is System.DBNull.Value Then
                            If dtErac.Rows(0).Item("ADDRESS_LINE_TWO") Is DBNull.Value Then
                                colParams.Add("<ERAC Address2>", dtErac.Rows(0).Item("CITY") + ", " + _
                                                                    dtErac.Rows(0).Item("STATE") + " " + _
                                                                    dtErac.Rows(0).Item("ZIP"))
                                colParams.Add("<ERAC City/State/Zip>", "")
                            Else
                                colParams.Add("<ERAC Address2>", dtErac.Rows(0).Item("ADDRESS_LINE_TWO"))
                                colParams.Add("<ERAC City/State/Zip>", dtErac.Rows(0).Item("CITY") + ", " + _
                                                                        dtErac.Rows(0).Item("STATE") + " " + _
                                                                        dtErac.Rows(0).Item("ZIP"))
                            End If
                        Else
                            colParams.Add("<ERAC Address2>", dtErac.Rows(0).Item("CITY") + ", " + _
                                                                dtErac.Rows(0).Item("STATE") + " " + _
                                                                dtErac.Rows(0).Item("ZIP"))
                            colParams.Add("<ERAC City/State/Zip>", "")
                        End If
                    ElseIf v.Count > 0 Then

                        If ERACPayee Then
                            strVendorName = IIf(strPayeeVendorName Is DBNull.Value, String.Empty, strPayeeVendorName)
                            strVendorCity = IIf(v.Item(0)("CITY") Is DBNull.Value, String.Empty, v.Item(0)("CITY"))
                            strVendorAddress1 = IIf(v.Item(0)("ADDRESS_ONE") Is DBNull.Value, String.Empty, v.Item(0)("ADDRESS_ONE"))
                            strVendorAddress2 = IIf(v.Item(0)("ADDRESS_TWO") Is DBNull.Value, String.Empty, v.Item(0)("ADDRESS_TWO"))

                            strVendorState = IIf(v.Item(0)("STATE") Is DBNull.Value, String.Empty, v.Item(0)("STATE"))
                            strVendorZip = IIf(v.Item(0)("ZIP") Is DBNull.Value, String.Empty, v.Item(0)("ZIP"))

                        End If


                        colParams.Add("<ERAC Address1>", v.Item(0)("ADDRESS_ONE"))

                        If v.Item(0)("ADDRESS_TWO") Is DBNull.Value OrElse v.Item(0)("ADDRESS_TWO") Is DBNull.Value Then
                            colParams.Add("<ERAC Address2>", v.Item(0)("CITY") + " " + _
                                                                v.Item(0)("STATE") + " " + _
                                                                v.Item(0)("ZIP"))
                            colParams.Add("<ERAC City/State/Zip>", "")
                        Else
                            colParams.Add("<ERAC Address2>", v.Item(0)("ADDRESS_TWO"))
                            colParams.Add("<ERAC City/State/Zip>", v.Item(0)("CITY") + ", " + _
                                                                    v.Item(0)("STATE") + " " + _
                                                                    v.Item(0)("ZIP"))
                        End If
                    Else
                        colParams.Add("<ERAC Contact>", String.Empty)
                        colParams.Add("<ERAC Address1>", String.Empty)
                        colParams.Add("<ERAC Address2>", String.Empty)
                        colParams.Add("<ERAC City/State/Zip>", String.Empty)
                    End If

                Else
                    'colParams.Add("<ERAC>", String.Empty)
                    colParams.Add("<ERAC Contact>", String.Empty)
                    colParams.Add("<ERAC Address1>", String.Empty)
                    colParams.Add("<ERAC Address2>", String.Empty)
                    colParams.Add("<ERAC City/State/Zip>", String.Empty)
                End If


            Else
                colParams.Add("<ERAC>", String.Empty)
                colParams.Add("<ERAC Contact>", String.Empty)
                colParams.Add("<ERAC Address1>", String.Empty)
                colParams.Add("<ERAC Address2>", String.Empty)
                colParams.Add("<ERAC City/State/Zip>", String.Empty)

            End If
            If oLustEvent.IRAC > 0 Then
                oIRACContact.Retrieve(oLustEvent.IRAC)
                colParams.Add("<IRAC>", oIRACContact.COMPANY_NAME)
            Else
                colParams.Add("<IRAC>", String.Empty)
            End If

            strVendorName = ""
            'Dim strContactName As String = ""
            If nFinEventID > 0 Then
                Dim drRow As DataRow
                Dim bolVendorPayee As Boolean = False

                oFinancialEvent.Retrieve(nFinEventID)
                Dim dtVendor As DataTable = GetXHAndXLContacts(nFinEventID, UIUtilsGen.EntityTypes.FinancialEvent, UIUtilsGen.ModuleID.Financial, , 1) 'oVendor.GetByID(oFinancialEvent.VendorID, nFinEventID, 616)
                Dim contactIDNotERAC As Boolean = False
                If dtVendor.Rows.Count > 0 Then
                    For Each drRow In dtVendor.Rows
                        If drRow.Item("Type") = EnumContactType.XH Then
                            If (Not drRow.Item("AssocCompany") Is System.DBNull.Value) AndAlso (drRow.Item("AssocCompany").ToString.Trim.Length <> 0) Then
                                If drRow.Item("IsPerson") = True Then
                                    strContactName = drRow.Item("CONTACT_Name")
                                    'colParams.Add("<ContactName>", drRow.Item("CONTACT_Name")) 'IIf(drRow.Item("Title") = String.Empty, "", drRow.Item("Title") + " ") + drRow.Item("First_Name") + " " + IIf(drRow.Item("Middle_Name") = String.Empty, "", drRow.Item("Middle_Name") + " ") + drRow.Item("Last_Name") + IIf(drRow.Item("Suffix") = String.Empty, "", " " + drRow.Item("Suffix")))
                                    'colParams.Add("<Salutation>", "Dear " & drRow.Item("CONTACT_Name")) 'IIf(drRow.Item("Title") = String.Empty, "", drRow.Item("Title") + " ") + drRow.Item("First_Name") + " " + IIf(drRow.Item("Middle_Name") = String.Empty, "", drRow.Item("Middle_Name") + " ") + drRow.Item("Last_Name") + IIf(drRow.Item("Suffix") = String.Empty, "", " " + drRow.Item("Suffix")))
                                End If
                                strVendorName = drRow.Item("AssocCompany")
                                strVendorID = IIf(drRow.Item("VENDOR_NUMBER") Is DBNull.Value, String.Empty, drRow.Item("VENDOR_NUMBER"))
                                ' contactIDNotERAC = True
                                contactIDNotERAC = False
                            Else
                                'If drRow.Item("IsPerson") = True Then
                                strVendorName = drRow.Item("CONTACT_Name") 'IIf(drRow.Item("Title") = String.Empty, "", drRow.Item("Title") + " ") + drRow.Item("First_Name") + " " + IIf(drRow.Item("Middle_Name") = String.Empty, "", drRow.Item("Middle_Name") + " ") + drRow.Item("Last_Name") + IIf(drRow.Item("Suffix") = String.Empty, "", " " + drRow.Item("Suffix"))
                                strVendorID = IIf(drRow.Item("VENDOR_NUMBER") Is DBNull.Value, String.Empty, drRow.Item("VENDOR_NUMBER"))
                                contactIDNotERAC = False 'True

                                'Else
                                '    strVendorName = drRow.Item("Company_Name")
                                'End If
                                'colParams.Add("<ContactName>", String.Empty)
                                'colParams.Add("<Salutation>", "Dear " & strVendorName)
                            End If
                            If drRow.Item("IsPerson") = True Then
                                bolIsPerson = True
                                nContactID = Integer.Parse(drRow.Item("ContactID"))
                            End If

                            If ERACPayee Then
                                cc = strVendorName

                                colParams.Add("<VendorERAC>", strPayeeVendorName)

                            Else
                                colParams.Add("<VendorERAC>", String.Empty)
                            End If

                            colParams.Add("<Vendor>", strVendorName)

                            If colParams.Item("<VendorAddress1>") Is Nothing Then

                                colParams.Add("<VendorAddress1>", IIf(drRow.Item("Address_One") Is DBNull.Value, "", drRow.Item("Address_One")))

                                If drRow.Item("Address_Two") Is DBNull.Value OrElse drRow.Item("Address_Two") = String.Empty Then
                                    colParams.Add("<VendorAddress2>", drRow.Item("City") & ", " & _
                                                                        drRow.Item("State") & " " & _
                                                                        drRow.Item("Zip"))
                                    colParams.Add("<VendorCityStateZip>", "")
                                Else
                                    colParams.Add("<VendorAddress2>", drRow.Item("Address_Two"))
                                    colParams.Add("<VendorCityStateZip>", drRow.Item("City") & ", " & _
                                                                            drRow.Item("State") & " " & _
                                                                            drRow.Item("Zip"))
                                End If
                            End If

                            If Not bolVendorPayee And Not ERACPayee Then
                                strPayeeVendorName = strVendorName
                                strVendorAddress1 = IIf(drRow.Item("Address_One") Is DBNull.Value, "", drRow.Item("Address_One"))
                                strVendorAddress2 = IIf(drRow.Item("Address_Two") Is DBNull.Value, "", drRow.Item("Address_Two"))
                                strVendorCity = IIf(drRow.Item("City") Is DBNull.Value, "", drRow.Item("City"))
                                strVendorState = IIf(drRow.Item("State") Is DBNull.Value, "", drRow.Item("State"))
                                strVendorZip = IIf(drRow.Item("Zip") Is DBNull.Value, "", drRow.Item("Zip"))
                            End If
                        ElseIf UCase(drRow.Item("ContactType")) = UCase("Financial Payee") AndAlso Not ERACPayee Then
                            bolVendorPayee = True
                            strPayeeVendorName = IIf(drRow.Item("AssocCompany") Is DBNull.Value OrElse drRow.Item("AssocCompany") = String.Empty, drRow.Item("CONTACT_Name"), drRow.Item("AssocCompany"))
                            strVendorID = IIf(drRow.Item("VENDOR_NUMBER") Is DBNull.Value, String.Empty, drRow.Item("VENDOR_NUMBER"))
                            strVendorAddress1 = IIf(drRow.Item("Address_One") Is DBNull.Value, "", drRow.Item("Address_One"))
                            strVendorAddress2 = IIf(drRow.Item("Address_Two") Is DBNull.Value, "", drRow.Item("Address_Two"))
                            strVendorCity = IIf(drRow.Item("City") Is DBNull.Value, "", drRow.Item("City"))
                            strVendorState = IIf(drRow.Item("State") Is DBNull.Value, "", drRow.Item("State"))
                            strVendorZip = IIf(drRow.Item("Zip") Is DBNull.Value, "", drRow.Item("Zip"))
                        ElseIf UCase(drRow.Item("ContactType")) = UCase("Financial Representative") Then
                            strFinRep = drRow.Item("CONTACT_Name")
                        End If

                        If drRow.Item("Type") = EnumContactType.XL Then
                            If drRow.Item("IsPerson") = True Then
                                bolIsPerson = True
                                nContactID = Integer.Parse(drRow.Item("ContactID"))
                                strContactName = drRow.Item("CONTACT_Name")
                                strContactFirstName = Trim(drRow.Item("First_Name"))
                            End If
                        End If

                    Next
                End If


                'Sets the ERAC Payee
                If ERACPayee Then

                    oContact.Retrieve(contactID)

                    If strVendorID Is Nothing Then
                        strVendorID = String.Empty
                        contactIDNotERAC = False
                    ElseIf contactIDNotERAC Then
                        contactIDNotERAC = False
                        strVendorID = String.Empty
                    End If

                    If v.Count > 0 AndAlso Not (v.Item(0)("Vendor_Number") Is DBNull.Value OrElse v.Item(0)("Vendor_Number") = "" OrElse v.Item(0)("Vendor_Number") = "0") Then
                        strVendorID = v.Item(0)("Vendor_Number")

                    ElseIf oContact Is Nothing OrElse contactID = 0 AndAlso strVendorID.Trim = String.Empty Then
                        strVendorID = "<Please Associate an ERAC contact with a vendor # and regenerate this notice>"
                    ElseIf Not oContact Is Nothing AndAlso Not oContact.VendorNumber Is Nothing AndAlso oContact.VendorNumber.Length > 0 Then
                        strVendorID = oContact.VendorNumber
                    ElseIf strVendorID.Trim = String.Empty Then
                        strVendorID = "<Please Associate an ERAC contact with a vendor # and regenerate this notice>"
                    End If
                ElseIf strVendorID Is Nothing OrElse strVendorID.Length = 0 Then
                    strVendorID = "<Please Associate an ERAC contact or financial vendor with a vendor #  and regenerate this notice>"
                End If

                colParams.Add("<PayeeVendor>", strPayeeVendorName)
                colParams.Add("<VendorNumber>", strVendorID)
                colParams.Add("<PayeeVendorAddress1>", strVendorAddress1)
                If strVendorAddress2 = String.Empty Then
                    colParams.Add("<PayeeVendorAddress2>", strVendorCity & IIf(strVendorCity.Length > 0, ", ", "") & _
                                                            strVendorState & " " & _
                                                            strVendorZip)
                    colParams.Add("<PayeeVendorCityStateZip>", "")
                Else
                    colParams.Add("<PayeeVendorAddress2>", strVendorAddress2)
                    colParams.Add("<PayeeVendorCityStateZip>", strVendorCity & IIf(strVendorCity.Length > 0, ", ", "") & _
                                                                strVendorState & " " & _
                                                                strVendorZip)
                End If

            Else
                If colParams.Item("<VendorAddress1>") Is Nothing Then
                    colParams.Add("<Vendor>", String.Empty)
                    colParams.Add("<VendorAddress1>", String.Empty)
                    colParams.Add("<VendorAddress2>", String.Empty)
                    colParams.Add("<VendorCityStateZip>", String.Empty)
                    colParams.Add("<Contact Name>", String.Empty)
                End If


            End If
            '------------------------Total Paid --- Added by Hua Cao  at Feb. 15, 2008----------
            Dim dtTotals As DataTable
            oFinancialEvent.Retrieve(nFinEventID)
            dtTotals = oFinancialEvent.CommitmentTotalsDatatable(0, False, False)
            colParams.Add("<TotalPaid>", dtTotals.Rows(0)("EventPaymentTotal").ToString())

            If ERACPayee Then


                colParams.Add("Vendor Address:", String.Empty)
                'colParams.Add("<,>", "<DeleteMe>" + "<DeleteMe>" + "<DeleteMe>")
            Else
                'colParams.Add("<,>", ",")


            End If

            '---------- salutation --------------------
            Dim dsSalutation As DataSet
            If bolIsPerson Then
                dsSalutation = oContact.GetContactLastName(nContactID)
            End If
            If strContactName <> String.Empty Then
                colParams.Add("<ContactName>", strContactName)
                If bolIsPerson Then

                    colParams.Add("<Salutation>", strContactName)
                    colParams.Add("<SalutationPage2>", strContactName)

                Else
                    colParams.Add("<Salutation>", strContactName)
                    colParams.Add("<SalutationPage2>", strContactName)
                End If

            ElseIf strVendorName <> String.Empty Then

                colParams.Add("<ContactName>", String.Empty)
                If bolIsPerson Then



                    colParams.Add("<Salutation>", strContactName)
                    colParams.Add("<SalutationPage2>", strContactName)


                Else
                    colParams.Add("<Salutation>", strContactName)
                    colParams.Add("<SalutationPage2>", strVendorName.Trim)
                End If
            Else
                If oOwnerInfo.OrganizationID > 0 Then
                    oPersonaInfo = pOwner.Organization
                    cc = pOwner.BPersona.Company
                    colParams.Add("<Vendor>", cc)

                    colParams.Add("<Salutation>", cc)
                Else
                    oPersonaInfo = pOwner.Persona
                    'colParams.Add("<VendorName>", pOwner.BPersona.FirstName.Trim & IIf(pOwner.BPersona.MiddleName.Trim.Length > 0, " ", "") & pOwner.BPersona.MiddleName.Trim & IIf(pOwner.BPersona.LastName.Trim.Length > 0, " ", "") & pOwner.BPersona.LastName.Trim & IIf(pOwner.BPersona.Suffix.Trim.Length > 0, " ", "") & pOwner.BPersona.Suffix.Trim)
                    cc = pOwner.BPersona.FirstName.Trim & _
                                                IIf(pOwner.BPersona.MiddleName.Trim.Length > 0, " ", "") & _
                                                pOwner.BPersona.MiddleName.Trim & _
                                                IIf(pOwner.BPersona.LastName.Trim.Length > 0, " ", "") & _
                                                pOwner.BPersona.LastName.Trim & _
                                                IIf(pOwner.BPersona.Suffix.Trim.Length > 0, " ", "") & _
                                                pOwner.BPersona.Suffix.Trim

                    colParams.Add("<Salutation>", cc)
                    colParams.Add("<SalutationPage2>", cc)
                End If

                colParams.Add("<Vendor>", cc)

                colParams.Add("<ContactName>", String.Empty)
            End If

            'oContact.GetAllByEntity(oOwnerInfo.ID, 614)

            'If oContact.ContactCollection.Count > 0 Then
            '    For Each oContactInfo In oContact.ContactCollection.Values
            '        If oContactInfo.orgCode > 0 Then
            '            strContactName = oContactInfo.companyName
            '        Else
            '            strContactName = oContactInfo.fullName
            '        End If
            '    Next
            'End If
            'colParams.Add("<ContactName>", strContactName)

            Dim strOwnerName As String = String.Empty
            If oOwnerInfo.OrganizationID > 0 Then
                oPersonaInfo = pOwner.Organization
                'colParams.Add("<Salutation>", pOwner.BPersona.Company.Trim)
                strOwnerName = pOwner.BPersona.Company
                colParams.Add("<OwnerName>", strOwnerName)
            Else
                oPersonaInfo = pOwner.Persona
                'colParams.Add("<Salutation>", "Dear " & pOwner.BPersona.Title.Trim & IIf(pOwner.BPersona.Title.Trim.Length > 0, " ", "") & pOwner.BPersona.FirstName.Trim & IIf(pOwner.BPersona.MiddleName.Trim.Length > 0, " ", "") & pOwner.BPersona.MiddleName.Trim & IIf(pOwner.BPersona.LastName.Trim.Length > 0, " ", "") & pOwner.BPersona.LastName.Trim & IIf(pOwner.BPersona.Suffix.Trim.Length > 0, " ", "") & pOwner.BPersona.Suffix.Trim & "")
                strOwnerName = pOwner.BPersona.FirstName.Trim & _
                                                IIf(pOwner.BPersona.MiddleName.Trim.Length > 0, " ", "") & _
                                                pOwner.BPersona.MiddleName.Trim & _
                                                IIf(pOwner.BPersona.LastName.Trim.Length > 0, " ", "") & _
                                                pOwner.BPersona.LastName.Trim & _
                                                IIf(pOwner.BPersona.Suffix.Trim.Length > 0, " ", "") & _
                                                pOwner.BPersona.Suffix.Trim
                colParams.Add("<OwnerName>", strOwnerName)
            End If

            If strFinRep = String.Empty AndAlso cc <> strOwnerName Then
                colParams.Add("<FinancialRep>", strOwnerName)
            Else
                colParams.Add("<FinancialRep>", strFinRep)
            End If

            If strVendorName = "" AndAlso colParams.Item("<VendorAddress1>") Is Nothing Then

                oAddressInfo = pOwner.Address()
                'colParams.Add("<Company Address>", oAddressInfo.AddressLine1)

                colParams.Add("<VendorAddress1>", oAddressInfo.AddressLine1)

                If oAddressInfo.AddressLine2 = String.Empty Then
                    colParams.Add("<VendorAddress2>", oAddressInfo.City & ", " & _
                                                        oAddressInfo.State.TrimEnd & " " & _
                                                        oAddressInfo.Zip)
                    colParams.Add("<VendorCityStateZip>", "")
                Else
                    colParams.Add("<VendorAddress2>", oAddressInfo.AddressLine2)
                    colParams.Add("<VendorCityStateZip>", oAddressInfo.City & ", " & _
                                                            oAddressInfo.State.TrimEnd & " " & _
                                                            oAddressInfo.Zip)
                End If
                'colParams.Add("<OwnerCityStateZip>", oAddressInfo.City & ", " & oAddressInfo.State.TrimEnd & " " & oAddressInfo.Zip)
            End If
            ownerAddressID = pOwner.AddressId
            oAddressInfo = pOwner.Addresses.Retrieve(oFacInfo.AddressID)
            pOwner.AddressId = ownerAddressID
            oOwnerInfo.AddressId = ownerAddressID
            colParams.Add("<FacilityName>", oFacInfo.Name)
            'colParams.Add("<FacilityStreetAddress>", oAddressInfo.AddressLine1 & " " & IIf(oAddressInfo.AddressLine2 = String.Empty, "", "; " & oAddressInfo.AddressLine2))
            colParams.Add("<Facility Address1>", oAddressInfo.AddressLine1)
            If oAddressInfo.AddressLine2 = String.Empty Then
                colParams.Add("<Facility Address2>", oAddressInfo.City & ", " & _
                                                        oAddressInfo.State.TrimEnd & " " & _
                                                        oAddressInfo.Zip)
                colParams.Add("<Facility City/State/Zip>", "")
            Else
                colParams.Add("<Facility Address2>", oAddressInfo.AddressLine2)
                colParams.Add("<Facility City/State/Zip>", oAddressInfo.City & ", " & _
                                                            oAddressInfo.State.TrimEnd & " " & _
                                                            oAddressInfo.Zip)
            End If

            'colParams.Add("<FacilityCityStateZip>", oAddressInfo.City.TrimEnd & ", " & oAddressInfo.State.TrimEnd & " " & oAddressInfo.Zip.TrimEnd)
            colParams.Add("<FacilityID>", oFacInfo.ID.ToString)



            colParams.Add("<ProjectManager>", oUserPM.Name)
            colParams.Add("<ProjectManagerPhone>", oUserPM.PhoneNumber)


            If nCommitmentID > 0 Then

                xCostFormat.AssignCommitmentObject(oCommitment)
                xCostFormat.CostFormatType = oCommitment.Case_Letter
                xCostFormat.SetDisplay(False)
                xCostFormat.LoadCommitment()
                sCommitmentGrandTotal = xCostFormat.GrandTotal
                Dim cF11To91TempStr As String
                Dim cF11T091ReplaceStr As String
                cF11To91TempStr = xCostFormat.lblCol1Row1.Text
                cF11T091ReplaceStr = cF11To91TempStr.Replace("Triannual", "Tri-annual")
                colParams.Add("<CF11>", cF11T091ReplaceStr)
                '   colParams.Add("<CF11>", xCostFormat.lblCol1Row1.Text)
                cF11To91TempStr = xCostFormat.lblCol1Row2.Text
                cF11T091ReplaceStr = cF11To91TempStr.Replace("Triannual", "Tri-annual")
                colParams.Add("<CF21>", cF11T091ReplaceStr)
                'colParams.Add("<CF21>", xCostFormat.lblCol1Row2.Text)
                cF11To91TempStr = xCostFormat.lblCol1Row3.Text
                cF11T091ReplaceStr = cF11To91TempStr.Replace("Triannual", "Tri-annual")
                colParams.Add("<CF31>", cF11T091ReplaceStr)
                'colParams.Add("<CF31>", xCostFormat.lblCol1Row3.Text)
                cF11To91TempStr = xCostFormat.lblCol1Row4.Text
                cF11T091ReplaceStr = cF11To91TempStr.Replace("Triannual", "Tri-annual")
                colParams.Add("<CF41>", cF11T091ReplaceStr)
                'colParams.Add("<CF41>", xCostFormat.lblCol1Row4.Text)
                cF11To91TempStr = xCostFormat.lblCol1Row5.Text
                cF11T091ReplaceStr = cF11To91TempStr.Replace("Triannual", "Tri-annual")
                colParams.Add("<CF51>", cF11T091ReplaceStr)
                'colParams.Add("<CF51>", xCostFormat.lblCol1Row5.Text)
                cF11To91TempStr = xCostFormat.lblCol1Row6.Text
                cF11T091ReplaceStr = cF11To91TempStr.Replace("Triannual", "Tri-annual")
                colParams.Add("<CF61>", cF11T091ReplaceStr)
                'colParams.Add("<CF61>", xCostFormat.lblCol1Row6.Text)
                cF11To91TempStr = xCostFormat.lblCol1Row7.Text
                cF11T091ReplaceStr = cF11To91TempStr.Replace("Triannual", "Tri-annual")
                colParams.Add("<CF71>", cF11T091ReplaceStr)
                'colParams.Add("<CF71>", xCostFormat.lblCol1Row7.Text)
                cF11To91TempStr = xCostFormat.lblCol1Row8.Text
                cF11T091ReplaceStr = cF11To91TempStr.Replace("Triannual", "Tri-annual")
                colParams.Add("<CF81>", cF11T091ReplaceStr)
                'colParams.Add("<CF81>", xCostFormat.lblCol1Row8.Text)
                cF11To91TempStr = xCostFormat.lblCol1Row9.Text
                cF11T091ReplaceStr = cF11To91TempStr.Replace("Triannual", "Tri-annual")
                colParams.Add("<CF91>", cF11T091ReplaceStr)
                'colParams.Add("<CF91>", xCostFormat.lblCol1Row9.Text)
                '--------- modified by Hua Cao 02/19/08 --------------------------------
                If xCostFormat.txtCol2Row1.Visible Then
                    Dim cF12TempStr As String
                    cF12TempStr = xCostFormat.txtCol2Row1.Text
                    colParams.Add("<CF12>", cF12TempStr)
                    '  colParams.Add("<CF12>", xCostFormat.txtCol2Row1.Text)
                Else
                    colParams.Add("<CF12>", String.Empty)
                End If
                If xCostFormat.txtCol2Row2.Visible Then
                    colParams.Add("<CF22>", xCostFormat.txtCol2Row2.Text)
                Else
                    colParams.Add("<CF22>", String.Empty)
                End If
                If xCostFormat.txtCol2Row3.Visible Then
                    colParams.Add("<CF32>", xCostFormat.txtCol2Row3.Text)
                Else
                    colParams.Add("<CF32>", String.Empty)
                End If
                If xCostFormat.txtCol2Row4.Visible Then
                    colParams.Add("<CF42>", xCostFormat.txtCol2Row4.Text)
                Else
                    colParams.Add("<CF42>", String.Empty)
                End If
                If xCostFormat.txtCol2Row5.Visible Then
                    colParams.Add("<CF52>", xCostFormat.txtCol2Row5.Text)
                Else
                    colParams.Add("<CF52>", String.Empty)
                End If
                If xCostFormat.txtCol2Row6.Visible Then
                    colParams.Add("<CF62>", xCostFormat.txtCol2Row6.Text)
                Else
                    colParams.Add("<CF62>", String.Empty)
                End If
                If xCostFormat.txtCol2Row7.Visible Then
                    colParams.Add("<CF72>", xCostFormat.txtCol2Row7.Text)
                Else
                    colParams.Add("<CF72>", String.Empty)
                End If
                If xCostFormat.txtCol2Row8.Visible Then
                    colParams.Add("<CF82>", xCostFormat.txtCol2Row8.Text)
                Else
                    colParams.Add("<CF82>", String.Empty)
                End If
                If xCostFormat.txtCol2Row9.Visible Then
                    colParams.Add("<CF92>", xCostFormat.txtCol2Row9.Text)
                Else
                    colParams.Add("<CF92>", String.Empty)
                End If


                colParams.Add("<CF13>", xCostFormat.lblCol3Row1.Text)
                colParams.Add("<CF23>", xCostFormat.lblCol3Row2.Text)
                colParams.Add("<CF33>", xCostFormat.lblCol3Row3.Text)
                colParams.Add("<CF43>", xCostFormat.lblCol3Row4.Text)
                colParams.Add("<CF53>", xCostFormat.lblCol3Row5.Text)
                colParams.Add("<CF63>", xCostFormat.lblCol3Row6.Text)
                colParams.Add("<CF73>", xCostFormat.lblCol3Row7.Text)
                colParams.Add("<CF83>", xCostFormat.lblCol3Row8.Text)
                colParams.Add("<CF93>", xCostFormat.lblCol3Row9.Text)

                If xCostFormat.txtCol4Row1.Visible Then
                    colParams.Add("<CF14>", xCostFormat.txtCol4Row1.Text)
                Else
                    colParams.Add("<CF14>", String.Empty)
                End If
                If xCostFormat.txtCol4Row2.Visible Then
                    colParams.Add("<CF24>", xCostFormat.txtCol4Row2.Text)
                Else
                    colParams.Add("<CF24>", String.Empty)
                End If
                If xCostFormat.txtCol4Row3.Visible Then
                    colParams.Add("<CF34>", xCostFormat.txtCol4Row3.Text)
                Else
                    colParams.Add("<CF34>", String.Empty)
                End If
                If xCostFormat.txtCol4Row4.Visible Then
                    colParams.Add("<CF44>", xCostFormat.txtCol4Row4.Text)
                Else
                    colParams.Add("<CF44>", String.Empty)
                End If
                If xCostFormat.txtCol4Row5.Visible Then
                    colParams.Add("<CF54>", xCostFormat.txtCol4Row5.Text)
                Else
                    colParams.Add("<CF54>", String.Empty)
                End If
                If xCostFormat.txtCol4Row6.Visible Then
                    colParams.Add("<CF64>", xCostFormat.txtCol4Row6.Text)
                Else
                    colParams.Add("<CF64>", String.Empty)
                End If
                If xCostFormat.txtCol4Row7.Visible Then
                    colParams.Add("<CF74>", xCostFormat.txtCol4Row7.Text)
                Else
                    colParams.Add("<CF74>", String.Empty)
                End If
                If xCostFormat.txtCol4Row8.Visible Then
                    colParams.Add("<CF84>", xCostFormat.txtCol4Row8.Text)
                Else
                    colParams.Add("<CF84>", String.Empty)
                End If
                If xCostFormat.txtCol4Row9.Visible Then
                    colParams.Add("<CF94>", xCostFormat.txtCol4Row9.Text)
                Else
                    colParams.Add("<CF94>", String.Empty)
                End If

                '------------Modified by Hua Cao  02/19/2008=-----------------------------
                Dim cF15To95TempStr As String
                If xCostFormat.lblCol5Row1.Text = "mo. =" Then
                    cF15To95TempStr = xCostFormat.lblCol5Row1.Text + "   "
                Else
                    cF15To95TempStr = xCostFormat.lblCol5Row1.Text
                End If
                colParams.Add("<CF15>", cF15To95TempStr)
                'colParams.Add("<CF15>", xCostFormat.lblCol5Row1.Text)
                If xCostFormat.lblCol5Row2.Text = "mo. =" Then
                    cF15To95TempStr = xCostFormat.lblCol5Row2.Text + "   "
                Else
                    cF15To95TempStr = xCostFormat.lblCol5Row2.Text
                End If
                colParams.Add("<CF25>", cF15To95TempStr)
                'colParams.Add("<CF25>", xCostFormat.lblCol5Row2.Text)
                If xCostFormat.lblCol5Row3.Text = "mo. =" Then
                    cF15To95TempStr = xCostFormat.lblCol5Row3.Text + "   "
                Else
                    cF15To95TempStr = xCostFormat.lblCol5Row3.Text
                End If
                colParams.Add("<CF35>", cF15To95TempStr)
                'colParams.Add("<CF35>", xCostFormat.lblCol5Row3.Text)
                If xCostFormat.lblCol5Row4.Text = "mo. =" Then
                    cF15To95TempStr = xCostFormat.lblCol5Row4.Text + "   "
                Else
                    cF15To95TempStr = xCostFormat.lblCol5Row4.Text
                End If
                colParams.Add("<CF45>", cF15To95TempStr)
                'colParams.Add("<CF45>", xCostFormat.lblCol5Row4.Text)
                If xCostFormat.lblCol5Row5.Text = "mo. =" Then
                    cF15To95TempStr = xCostFormat.lblCol5Row5.Text + "   "
                Else
                    cF15To95TempStr = xCostFormat.lblCol5Row5.Text
                End If
                colParams.Add("<CF55>", cF15To95TempStr)
                'colParams.Add("<CF55>", xCostFormat.lblCol5Row5.Text)
                If xCostFormat.lblCol5Row6.Text = "mo. =" Then
                    cF15To95TempStr = xCostFormat.lblCol5Row6.Text + "   "
                Else
                    cF15To95TempStr = xCostFormat.lblCol5Row6.Text
                End If
                colParams.Add("<CF65>", cF15To95TempStr)
                'colParams.Add("<CF65>", xCostFormat.lblCol5Row6.Text)
                If xCostFormat.lblCol5Row7.Text = "mo. =" Then
                    cF15To95TempStr = xCostFormat.lblCol5Row7.Text + "   "
                Else
                    cF15To95TempStr = xCostFormat.lblCol5Row7.Text
                End If
                colParams.Add("<CF75>", cF15To95TempStr)
                'colParams.Add("<CF75>", xCostFormat.lblCol5Row7.Text)
                If xCostFormat.lblCol5Row8.Text = "mo. =" Then
                    cF15To95TempStr = xCostFormat.lblCol5Row8.Text + "   "
                Else
                    cF15To95TempStr = xCostFormat.lblCol5Row8.Text
                End If
                colParams.Add("<CF85>", cF15To95TempStr)
                'colParams.Add("<CF85>", xCostFormat.lblCol5Row8.Text)
                If xCostFormat.lblCol5Row9.Text = "mo. =" Then
                    cF15To95TempStr = xCostFormat.lblCol5Row9.Text + "   "
                Else
                    cF15To95TempStr = xCostFormat.lblCol5Row9.Text
                End If
                colParams.Add("<CF95>", cF15To95TempStr)
                'colParams.Add("<CF95>", xCostFormat.lblCol5Row9.Text)

                If xCostFormat.txtCol6Row1.Visible Then
                    Dim cF16TempStr As String
                    cF16TempStr = xCostFormat.txtCol6Row1.Text
                    colParams.Add("<CF16>", cF16TempStr)
                    '   colParams.Add("<CF16>", xCostFormat.txtCol6Row1.Text)
                Else
                    colParams.Add("<CF16>", String.Empty)
                End If
                If xCostFormat.txtCol6Row2.Visible Then
                    colParams.Add("<CF26>", xCostFormat.txtCol6Row2.Text)
                Else
                    colParams.Add("<CF26>", String.Empty)
                End If
                If xCostFormat.txtCol6Row3.Visible Then
                    colParams.Add("<CF36>", xCostFormat.txtCol6Row3.Text)
                Else
                    colParams.Add("<CF36>", String.Empty)
                End If
                If xCostFormat.txtCol6Row4.Visible Then
                    colParams.Add("<CF46>", xCostFormat.txtCol6Row4.Text)
                Else
                    colParams.Add("<CF46>", String.Empty)
                End If
                If xCostFormat.txtCol6Row5.Visible Then
                    colParams.Add("<CF56>", xCostFormat.txtCol6Row5.Text)
                Else
                    colParams.Add("<CF56>", String.Empty)
                End If
                If xCostFormat.txtCol6Row6.Visible Then
                    colParams.Add("<CF66>", xCostFormat.txtCol6Row6.Text)
                Else
                    colParams.Add("<CF66>", String.Empty)
                End If
                If xCostFormat.txtCol6Row7.Visible Then
                    colParams.Add("<CF76>", xCostFormat.txtCol6Row7.Text)
                Else
                    colParams.Add("<CF76>", String.Empty)
                End If
                If xCostFormat.txtCol6Row8.Visible Then
                    colParams.Add("<CF86>", xCostFormat.txtCol6Row8.Text)
                Else
                    colParams.Add("<CF86>", String.Empty)
                End If
                If xCostFormat.txtCol6Row9.Visible Then
                    colParams.Add("<CF96>", xCostFormat.txtCol6Row9.Text)
                Else
                    colParams.Add("<CF96>", String.Empty)
                End If

                Dim dueStateStmStr As String
                '-------- Modified by Hua Cao 02/20/2008 ---------------------
                dueStateStmStr = oCommitment.DueDateStatement
                colParams.Add("<DueDateText>", dueStateStmStr.Replace("Triannual", "Tri-annual"))
                '  colParams.Add("<DueDateText>", oCommitment.DueDateStatement)
                colParams.Add("<CommitmentAmount>", FormatNumber(sCommitmentGrandTotal, 2, TriState.True, TriState.UseDefault, TriState.True))
                colParams.Add("<PurchaseOrderNumber>", oCommitment.PONumber)
                colParams.Add("<SOW DATE>", Format(oCommitment.SOWDate, "MMMM d, yyyy"))
                colParams.Add("<Approved Date>", Format(oCommitment.ApprovedDate, "MMMM d, yyyy"))

                oActivity.Retrieve(CInt(oCommitment.ActivityType))
                'colParams.Add("<SOW>", oProperty.PropDesc)
                'added by Hua Cao     Feb. 15 2008--------------------------------
                Dim tempSOWStr As String
                Dim tempYear As String
                Dim strSOWYear As String
                tempSOWStr = oActivity.ActivityDesc

                tempYear = tempSOWStr.Substring(tempSOWStr.Length() - 1)
                Try
                    If (CInt(tempYear) < 9999) Then
                        If (tempSOWStr.Substring(0, 11) = "Groundwater") Then
                            strSOWYear = tempSOWStr.Insert(tempSOWStr.Length(), " Event(s)")
                        Else
                            strSOWYear = tempSOWStr.Insert(tempSOWStr.Length() - 1, "   Year ")
                        End If
                        colParams.Add("<SOW>", strSOWYear)
                    Else
                        colParams.Add("<SOW>", oActivity.ActivityDesc)
                    End If
                Catch Ex As Exception
                    colParams.Add("<SOW>", oActivity.ActivityDesc)
                End Try







                ' colParams.Add("<SOW>", oActivity.ActivityDesc)

                oProperty.Retrieve(CInt(oCommitment.ContractType))
                'colParams.Add("<ReimbursementType>", oProperty.PropDesc)
                colParams.Add("<ReimbursementType>", oProperty.Name)

            Else

                colParams.Add("<CF11>", String.Empty)
                colParams.Add("<CF21>", String.Empty)
                colParams.Add("<CF31>", String.Empty)
                colParams.Add("<CF41>", String.Empty)
                colParams.Add("<CF51>", String.Empty)
                colParams.Add("<CF61>", String.Empty)
                colParams.Add("<CF71>", String.Empty)
                colParams.Add("<CF81>", String.Empty)
                colParams.Add("<CF91>", String.Empty)

                colParams.Add("<CF12>", String.Empty)
                colParams.Add("<CF22>", String.Empty)
                colParams.Add("<CF32>", String.Empty)
                colParams.Add("<CF42>", String.Empty)
                colParams.Add("<CF52>", String.Empty)
                colParams.Add("<CF62>", String.Empty)
                colParams.Add("<CF72>", String.Empty)
                colParams.Add("<CF82>", String.Empty)
                colParams.Add("<CF92>", String.Empty)

                colParams.Add("<CF13>", String.Empty)
                colParams.Add("<CF23>", String.Empty)
                colParams.Add("<CF33>", String.Empty)
                colParams.Add("<CF43>", String.Empty)
                colParams.Add("<CF53>", String.Empty)
                colParams.Add("<CF63>", String.Empty)
                colParams.Add("<CF73>", String.Empty)
                colParams.Add("<CF83>", String.Empty)
                colParams.Add("<CF93>", String.Empty)

                colParams.Add("<CF14>", String.Empty)
                colParams.Add("<CF24>", String.Empty)
                colParams.Add("<CF34>", String.Empty)
                colParams.Add("<CF44>", String.Empty)
                colParams.Add("<CF54>", String.Empty)
                colParams.Add("<CF64>", String.Empty)
                colParams.Add("<CF74>", String.Empty)
                colParams.Add("<CF84>", String.Empty)
                colParams.Add("<CF94>", String.Empty)

                colParams.Add("<CF15>", String.Empty)
                colParams.Add("<CF25>", String.Empty)
                colParams.Add("<CF35>", String.Empty)
                colParams.Add("<CF45>", String.Empty)
                colParams.Add("<CF55>", String.Empty)
                colParams.Add("<CF65>", String.Empty)
                colParams.Add("<CF75>", String.Empty)
                colParams.Add("<CF85>", String.Empty)
                colParams.Add("<CF95>", String.Empty)

                colParams.Add("<CF16>", String.Empty)
                colParams.Add("<CF26>", String.Empty)
                colParams.Add("<CF36>", String.Empty)
                colParams.Add("<CF46>", String.Empty)
                colParams.Add("<CF56>", String.Empty)
                colParams.Add("<CF66>", String.Empty)
                colParams.Add("<CF76>", String.Empty)
                colParams.Add("<CF86>", String.Empty)
                colParams.Add("<CF96>", String.Empty)

                colParams.Add("<DueDateText>", String.Empty)
                colParams.Add("<CommitmentAmount>", String.Empty)
                colParams.Add("<PurchaseOrderNumber>", String.Empty)
                colParams.Add("<SOW DATE>", String.Empty)
                colParams.Add("<Approved Date>", Format(Now, "MMMM d, yyyy"))
                colParams.Add("<SOW>", String.Empty)
                colParams.Add("<ReimbursementType>", String.Empty)


            End If

            If nAdjustmentID > 0 Then
                oAdjustment.Retrieve(nAdjustmentID)
                colParams.Add("<ChangeOrderAmount>", FormatNumber(oAdjustment.AdjustAmount, 2, TriState.True, TriState.False, TriState.True))
            Else
                colParams.Add("<ChangeOrderAmount>", String.Empty)
            End If

            If nReimbursementID > 0 Then

                If oReimbursement.Incomplete Then
                    Dim dtTable As DataTable
                    Dim dRow As DataRow
                    Dim strReasons As String
                    dtTable = oReimbursement.GetFinancialIncompleteAppReasonsForLetters(oReimbursement.IncompleteReason)
                    For Each dRow In dtTable.Rows
                        'strReasons += Trim(dRow("Financial_Text")) + vbCrLf
                        If CInt(dRow("Text_ID")) = 115 Then
                            strReasons += Trim(oReimbursement.IncompleteOther) + vbCrLf
                        Else
                            strReasons += Trim(dRow("Financial_Text")) + vbCrLf
                        End If
                    Next

                    colParams.Add("<Reasons>", strReasons)
                    'xArray = oReimbursement.IncompleteReason.Split(",")
                    'Dim strColParamKey As String = String.Empty

                    'For i = 0 To xArray.Length - 1
                    '    Select Case xArray(i)

                    '        Case 103
                    '            colParams.Add("<IA1>", "X")
                    '            strColParamKey += "<IA1>" + ","
                    '        Case 104
                    '            colParams.Add("<IA2>", "X")
                    '            strColParamKey += "<IA2>" + ","
                    '        Case 105
                    '            colParams.Add("<IA3>", "X")
                    '            strColParamKey += "<IA3>" + ","
                    '        Case 106
                    '            colParams.Add("<IA4>", "X")
                    '            strColParamKey += "<IA4>" + ","
                    '        Case 107
                    '            colParams.Add("<IA5>", "X")
                    '            strColParamKey += "<IA5>" + ","
                    '        Case 108
                    '            colParams.Add("<IA6>", "X")
                    '            strColParamKey += "<IA6>" + ","
                    '        Case 109
                    '            colParams.Add("<IA7>", "X")
                    '            strColParamKey += "<IA7>" + ","
                    '        Case 110
                    '            colParams.Add("<IA8>", "X")
                    '            strColParamKey += "<IA8>" + ","
                    '        Case 111
                    '            colParams.Add("<IA9>", "X")
                    '            strColParamKey += "<IA9>" + ","
                    '        Case 112
                    '            colParams.Add("<IA10>", "X")
                    '            strColParamKey += "<IA10>" + ","
                    '        Case 113
                    '            colParams.Add("<IA11>", "X")
                    '            strColParamKey += "<IA11>" + ","
                    '        Case 114
                    '            colParams.Add("<IA12>", "X")
                    '            strColParamKey += "<IA12>" + ","
                    '        Case 115
                    '            colParams.Add("<IA13>", "X")
                    '            colParams.Add("<IAOther>", oReimbursement.IncompleteOther)
                    '            strColParamKey += "<IA13>" + "," + "<IA14>"
                    '    End Select

                    'Next
                    'If strColParamKey.EndsWith(",") Then
                    '    strColParamKey = strColParamKey.Substring(0, strColParamKey.Length - 1)
                    'End If
                    'Dim strKeys() As String = strColParamKey.Split(",")
                    'Dim bolKeyExists As Boolean = False
                    'For j As Integer = 1 To 14
                    '    For k As Integer = 0 To UBound(strKeys)
                    '        If strKeys(k).IndexOf("<IA" + j.ToString + ">") >= 0 Then
                    '            bolKeyExists = True
                    '            Exit For
                    '        End If
                    '    Next
                    '    If Not bolKeyExists Then
                    '        If j = 14 Then
                    '            colParams.Add("<IAOther>", String.Empty)
                    '        Else
                    '            colParams.Add("<IA" + j.ToString + ">", String.Empty)
                    '        End If
                    '    End If
                    '    bolKeyExists = False
                    'Next

                Else
                    'colParams.Add("<IA1>", String.Empty)
                    'colParams.Add("<IA2>", String.Empty)
                    'colParams.Add("<IA3>", String.Empty)
                    'colParams.Add("<IA4>", String.Empty)
                    'colParams.Add("<IA5>", String.Empty)
                    'colParams.Add("<IA6>", String.Empty)
                    'colParams.Add("<IA7>", String.Empty)
                    'colParams.Add("<IA8>", String.Empty)
                    'colParams.Add("<IA9>", String.Empty)
                    'colParams.Add("<IA10>", String.Empty)
                    'colParams.Add("<IA11>", String.Empty)
                    'colParams.Add("<IA12>", String.Empty)
                    'colParams.Add("<IA13>", String.Empty)
                    'colParams.Add("<IAOther>", String.Empty)

                End If

                colParams.Add("<PaymentNumber>", oReimbursement.PaymentNumber)
                colParams.Add("<Received Date>", oReimbursement.ReceivedDate.ToShortDateString)

                strInvoiceNumbers = ""
                strInvoiceComments = ""
                strDeductionReasons = ""
                nPaidAmount = 0
                nRequestAmount = 0







                colInvoice = oInvoice.GetAllByReimbursement(nReimbursementID)
                For Each oInvoiceInfo In colInvoice.Values
                    If strInvoiceNumbers <> "" Then
                        strInvoiceNumbers &= ", "
                    End If

                    strInvoiceNumbers &= oInvoiceInfo.VendorInvoice
                    strDeductionReasons &= IIf(oInvoiceInfo.DeductionReason = String.Empty, String.Empty, oInvoiceInfo.DeductionReason & vbCrLf)
                    nPaidAmount += oInvoiceInfo.PaidAmount
                    nRequestAmount += oInvoiceInfo.InvoicedAmount

                    If oInvoiceInfo.Comment <> String.Empty Then
                        If strInvoiceComments <> String.Empty Then
                            strInvoiceComments += ", "
                        End If
                        strInvoiceComments += oInvoiceInfo.Comment
                    End If
                Next
                colParams.Add("<INVOICE DATE>", Today.ToShortDateString)
                If nPaidAmount = nRequestAmount Then
                    colParams.Add("<X1>", "X")
                    colParams.Add("<X2>", String.Empty)
                    colParams.Add("<REQUEST AMOUNT>", "$" + CStr(FormatNumber(nRequestAmount, 2, TriState.True, TriState.False, TriState.True)))
                    ' colParams.Add("<REQUEST AMOUNT2>", String.Empty)
                    colParams.Add("<REQUEST AMOUNT2>", "__________")
                    colParams.Add("<PAID AMOUNT>", "__________")
                Else
                    colParams.Add("<X1>", String.Empty)
                    colParams.Add("<X2>", "X")
                    colParams.Add("<REQUEST AMOUNT2>", "$" + CStr(FormatNumber(nRequestAmount, 2, TriState.True, TriState.False, TriState.True)))
                    'colParams.Add("<REQUEST AMOUNT>", String.Empty)
                    colParams.Add("<REQUEST AMOUNT>", "__________")
                    colParams.Add("<PAID AMOUNT>", "$" + CStr(FormatNumber(nPaidAmount, 2, TriState.True, TriState.False, TriState.True)))
                End If
                colParams.Add("<ApprovedAmount>", CStr(FormatNumber(nPaidAmount, 2, TriState.True, TriState.False, TriState.True)))
                'colParams.Add("<PAID AMOUNT>", "$" + CStr(FormatNumber(nPaidAmount, 2, TriState.True, TriState.False, TriState.True)))
                ' colParams.Add("<DEDUCTION REASONS>", strDeductionReasons)

                colParams.Add("<INVOICE NUMBERS>", strInvoiceNumbers)
                colParams.Add("<Comments>", strInvoiceComments)
                colParams.Add("<UnencumberAmount>", CStr(FormatNumber((nRequestAmount - nPaidAmount), 2, TriState.True, TriState.False, TriState.True)))

            Else
                'colParams.Add("<IA1>", String.Empty)
                'colParams.Add("<IA2>", String.Empty)
                'colParams.Add("<IA3>", String.Empty)
                'colParams.Add("<IA4>", String.Empty)
                'colParams.Add("<IA5>", String.Empty)
                'colParams.Add("<IA6>", String.Empty)
                'colParams.Add("<IA7>", String.Empty)
                'colParams.Add("<IA8>", String.Empty)
                'colParams.Add("<IA9>", String.Empty)
                'colParams.Add("<IA10>", String.Empty)
                'colParams.Add("<IA11>", String.Empty)
                'colParams.Add("<IA12>", String.Empty)
                'colParams.Add("<IA13>", String.Empty)
                'colParams.Add("<IAOther>", String.Empty)


                colParams.Add("<PaymentNumber>", String.Empty)
                colParams.Add("<INVOICE DATE>", Today.ToShortDateString)
                colParams.Add("<X1>", String.Empty)
                colParams.Add("<X2>", String.Empty)
                colParams.Add("<REQUEST AMOUNT2>", String.Empty)
                colParams.Add("<REQUEST AMOUNT>", String.Empty)
                colParams.Add("<ApprovedAmount>", String.Empty)
                colParams.Add("<PAID AMOUNT>", String.Empty)
                'colParams.Add("<DEDUCTION REASONS>", String.Empty)
                colParams.Add("<INVOICE NUMBERS>", String.Empty)
                colParams.Add("<Comments>", String.Empty)
                colParams.Add("<UnencumberAmount>", String.Empty)
            End If

            oUserExecutiveDirectorInfo = oUserPM.RetrieveExecutiveDirector()
            colParams.Add("<Executive Director>", oUserExecutiveDirectorInfo.Name)
            colParams.Add("<Executive Director Title case>", UIUtilsGen.TitleCaseString(oUserExecutiveDirectorInfo.Name))
            colParams.Add("<Executive Director CAPS>", oUserExecutiveDirectorInfo.Name.ToUpper)

            If nInvoiceID > 0 Then
                oInvoice.Retrieve(nInvoiceID)
                colParams.Add("<InvoicePaidAmount>", FormatNumber(oInvoice.PaidAmount, 2, TriState.True, TriState.False, TriState.True))
            Else
                colParams.Add("<InvoicePaidAmount>", String.Empty)
            End If
            If Not ERACPayee Then
                FillCCList(nFinEventID, UIUtilsGen.ModuleID.Financial, colParams)
            Else
                colParams.Add("<CC list>", String.Format(" CC: {0}", cc))
            End If


            If (Not pOwner Is Nothing AndAlso Not pOwner.Organization Is Nothing) And (colParams.Item("<Vendor>") Is Nothing OrElse colParams.Item("<Vendor>") = String.Empty) Then
                colParams.Remove("<Vendor>")
                colParams.Add("<Vendor>", pOwner.Organization.Company)
                colParams.Remove("<VendorERAC>")
                colParams.Add("<VendorERAC>", strPayeeVendorName)

            End If

            If colParams.Item("<VendorCityStateZip>") Is Nothing Then
                colParams.Add("<VendorCityStateZip>", String.Empty)
            End If

            Try
                Dim strTempPath As String = TmpltPath & "Financial\" & strTemplateName
                If nFinEventID > 0 Then
                    nOwningEntity = nFinEventID
                    nEntityID = UIUtilsGen.EntityTypes.FinancialEvent
                End If
                If nCommitmentID > 0 Then
                    nOwningEntity = nCommitmentID
                    nEntityID = UIUtilsGen.EntityTypes.FinancialCommitment
                End If
                If nReimbursementID > 0 Then
                    nOwningEntity = nReimbursementID
                    nEntityID = UIUtilsGen.EntityTypes.FinancialReimbursement
                End If
                If Not UCase(strTempPath).StartsWith(UCase(TmpltPath & "Financial\" & "MGPTFApprovalFormTemplate")) And strTempPath <> TmpltPath & "Financial\" & "NoticeofReimbursementTemplate.doc" Then
                    UIUtilsGen.CreateAndSaveDocument("Financial", nOwningEntity, nEntityID, DOC_PATH, strDOC_NAME, strDocType, strTempPath, Doc_Path, strDocDesc, nModuleID, colParams, eventID, eventSequence, eventType)
                Else
                    If UCase(strTempPath) = UCase(TmpltPath & "Financial\" & "NoticeofReimbursementTemplate.doc") Then
                        If nReimbursementID > 0 Then
                            colParams.Add("<DEDUCTION REASONS>", strDeductionReasons)
                        Else
                            colParams.Add("<DEDUCTION REASONS>", String.Empty)
                        End If
                    End If
                    If UCase(strTempPath).StartsWith(UCase(TmpltPath & "Financial\" & "MGPTFApprovalFormTemplate")) Then
                        If nCommitmentID > 0 Then
                            colParams.Add("<Reimbursement Conditions>", oCommitment.ReimbursementCondition.Trim)
                        Else
                            colParams.Add("<Reimbursement Conditions>", String.Empty)
                        End If
                    End If

                    Dim oWord As Word.Application = MusterContainer.GetWordApp

                    If Not oWord Is Nothing Then

                        Dim ltrGen As New MUSTER.BusinessLogic.pLetterGen
                        ltrGen.CreateFinancialLetter("Financial", strDOC_NAME, colParams, strTempPath, Doc_Path & strDOC_NAME, oWord)
                        UIUtilsGen.SaveDocument(nOwningEntity, nEntityID, strDOC_NAME, strDocType, Doc_Path, strDocDesc, nModuleID, eventID, eventSequence, eventType)
                        ltrGen = Nothing
                    End If
                    oWord = Nothing
                End If


                v = Nothing
            Catch ex As Exception
                Throw ex

            End Try

        Catch ex As Exception
            Throw ex
        Finally
            SetCursorType(System.Windows.Forms.Cursors.Default)
        End Try
    End Function
    Friend Function GenerateFinancialGenericLetter(ByVal ModuleID As Integer, ByVal Title As String, ByVal BodyData As String, ByVal Cols As Int16, Optional ByRef IsDraft As Boolean = True, Optional ByVal ModuleName As String = "Global", Optional ByRef EntityID As Int64 = 0, Optional ByVal EntityType As Int64 = 0, Optional ByVal DocNamePrefix As String = "FAC_GEN_", Optional ByVal DocDescription As String = "", Optional ByVal DocType As String = "", Optional ByVal TableFormat As Int16 = 35, Optional ByVal bApplyBorders As Boolean = False, Optional ByVal bApplyShading As Boolean = False, Optional ByVal bApplyFont As Boolean = True, Optional ByVal bApplyColor As Boolean = False, Optional ByVal bApplyHeader As Boolean = True, Optional ByVal bApplyLastRow As Boolean = False, Optional ByVal bApplyFirstCol As Boolean = False, Optional ByVal bApplyLastCol As Boolean = False, Optional ByVal bApplyAutofit As Boolean = True, Optional ByVal eventID As Int64 = 0, Optional ByVal eventSequence As Integer = 0, Optional ByVal eventType As Integer = 0) As Boolean

        Dim bolNeedCapDocs As Boolean = False
        Dim strDOC_NAME As String = String.Empty
        Dim strToday As String = String.Empty
        Dim i As Int16
        Dim strYes As String
        Dim strNo As String
        Dim strOther As String
        Dim colParams As New Specialized.NameValueCollection
        Dim strDocPath As String
        Dim tmpDate As Date
        Dim fileName As Object
        Dim aDoc As Word.Document
        Dim areadOnly As Object = True
        Dim isVisible As Object = True
        Dim confirmConversions As Object = False
        Dim addToRecentFiles As Object = False
        Dim revert As Object = False
        Dim missing As Object = System.Reflection.Missing.Value

        Try

            If DOC_PATH = "\" Then
                Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
            End If

            strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm")) + "_" + CStr(Format(Now, "ss"))
            strDOC_NAME = DocNamePrefix + CStr(Trim(EntityID.ToString)) + "_" + strToday + ".doc"
            colParams.Add("<DATE>", Format(Now, "MMMM d, yyyy"))
            colParams.Add("<TITLE>", Title)
            colParams.Add("<DATA>", BodyData)

            Dim ltrGen As New MUSTER.BusinessLogic.pLetterGen
            Dim strTempPath As String = System.IO.Path.GetTempPath

            If IsDraft Then
                strDocPath = strTempPath
            Else
                strDocPath = DOC_PATH
            End If

            Dim oWordApp As Word.Application = MusterContainer.GetWordApp

            If Not oWordApp Is Nothing Then


                ltrGen.CreateFinancialGenericLetter(ModuleName, strDOC_NAME, colParams, Cols, TmpltPath & "Global\Generic.doc", strDocPath & strDOC_NAME, oWordApp, TableFormat, bApplyBorders, bApplyShading, bApplyFont, bApplyColor, bApplyHeader, bApplyLastRow, bApplyFirstCol, bApplyLastCol, bApplyAutofit)
                'Delay()
                UIUtilsGen.Delay(, 1)


                ' Make word visible
                oWordApp.Visible = True

                If IsDraft Then
                    ' Open the document that was chosen by the dialog
                    aDoc = oWordApp.Documents.Open(strDocPath & strDOC_NAME, confirmConversions, areadOnly, addToRecentFiles, missing, missing, revert, missing, missing, missing, missing, isVisible)

                Else
                    UIUtilsGen.SaveDocument(EntityID, EntityType, strDOC_NAME, DocType, DOC_PATH, DocDescription, ModuleID, eventID, eventSequence, eventType)
                End If
            End If
            oWordApp = Nothing



        Catch ex As Exception
            Throw ex
        Finally
            SetCursorType(System.Windows.Forms.Cursors.Default)
        End Try
    End Function
#End Region
#Region "Company Letters"
    Friend Function GenerateLicenseeLetter(ByVal nEntityID As String, ByVal strDocType As String, ByVal strDocName As String, ByVal strDocDesc As String, ByVal strTemplateName As String, Optional ByVal Params As Specialized.NameValueCollection = Nothing, Optional ByVal strFile As String = "", Optional ByVal Selrow As Infragistics.Win.UltraWinGrid.UltraGridRow = Nothing, Optional ByVal strUserId As String = "", Optional ByVal strSignature As String = "") As Boolean
        Dim strDOC_NAME As String = String.Empty
        Dim strToday As String = String.Empty
        Dim colParams As New Specialized.NameValueCollection
        Dim pLicen As New MUSTER.BusinessLogic.pLicensee
        Dim strInfoNeeded As String = String.Empty
        Dim count As Integer = 0
        nModuleID = UIUtilsGen.ModuleID.Company
        Try

            If DOC_PATH = "\" Then
                Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
            End If

            strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm"))
            strDOC_NAME = "COM_" + strDocName.Trim.ToString + "_" + CStr(Trim(nEntityID.ToString)) + "_" + strToday + ".doc"

            Dim ServerPath As String = TmpltPath + "\Company"
            Dim Localpath As String = "C:\Templates\Company".Trim
            CopyTemplatesFromServerToLocal(ServerPath, Localpath)
            Try
                If Params Is Nothing Then

                    colParams.Add("<Date>", Format(Now, "MMMM d, yyyy"))
                    'colParams.Add("<Expiration Date>", Selrow.Cells("Expire_Date").Value.ToShortDateString)
                    If Not Selrow.Cells("Expire_Date").Text = String.Empty Then
                        colParams.Add("<Expiration Date>", Selrow.Cells("Expire_Date").Value.ToShortDateString)
                        'ElseIf Not Selrow.Cells("Expire_Date").Text Is Nothing Then
                        '    colParams.Add("<Expiration Date>", Selrow.Cells("Expire_Date").Value.ToShortDateString)
                    Else
                        colParams.Add("<Expiration Date>", String.Empty)
                    End If
                    colParams.Add("<Licensee Name>", Selrow.Cells("Licensee").Value)
                    colParams.Add("<Company Name>", Selrow.Cells("CompanyName").Value)
                    colParams.Add("<Company Address1>", Selrow.Cells("Address1").Value)


                    Dim strZip As String = Selrow.Cells("Zip").Value
                    strZip = strZip.Substring(0, 5)
                    'If Selrow.Cells("Address2").Value Is System.DBNull.Value Then
                    If Selrow.Cells("Address2").Text = String.Empty Then
                        'colParams.Add("<Company Address2>", Selrow.Cells("City").Value & ", " & Selrow.Cells("State").Value.TrimEnd & " " & strZip)
                        colParams.Add("<Company Address2>", IIf(Selrow.Cells("City").Value Is System.DBNull.Value, "", Selrow.Cells("City").Value & ", ") & IIf(Selrow.Cells("State").Value Is System.DBNull.Value, "", Selrow.Cells("State").Value & " ") & strZip)
                        colParams.Add("<City/State/Zip>", "")
                    Else
                        colParams.Add("<Company Address2>", Selrow.Cells("Address2").Value)
                        colParams.Add("<City/State/Zip>", IIf(Selrow.Cells("City").Value Is System.DBNull.Value, "", Selrow.Cells("City").Value & ", ") & IIf(Selrow.Cells("State").Value Is System.DBNull.Value, "", Selrow.Cells("State").Value & " ") & strZip)
                    End If
                    If Selrow.Cells("Cert_Type").Value.ToString.ToUpper = "CLOSURE" Then
                        colParams.Add("<Certification Type>", "Permanently Close")

                    ElseIf Selrow.Cells("Cert_Type").Value.ToString.ToUpper = "INSTALL" Then
                        colParams.Add("<Certification Type>", "Install, Alter, and Permanently Close")
                    End If
                    'If Not Selrow.Cells("App_recvd_date").Value Is System.DBNull.Value Then

                    If Selrow.Cells("LicenseeNo").Text.StartsWith("CHB") Or _
                        Selrow.Cells("LicenseeNo").Text.StartsWith("CRX") Or _
                        Selrow.Cells("LicenseeNo").Text.StartsWith("CHX") Or _
                        Selrow.Cells("LicenseeNo").Text.StartsWith("NHB") Or _
                        Selrow.Cells("LicenseeNo").Text.StartsWith("NRX") Or _
                        Selrow.Cells("LicenseeNo").Text.StartsWith("NHX") Then
                        If Not Selrow.Cells("App_recvd_date").Text = String.Empty Then
                            If Not Selrow.Cells("Issued_date").Text = String.Empty Then
                                If CDate(Selrow.Cells("App_recvd_date").Text) < CDate(Selrow.Cells("Issued_date").Text) Then
                                    count += 1
                                    '    strInfoNeeded += count.ToString + ". A completed certification renewal application" + vbCrLf
                                End If
                            End If
                        End If
                    End If

                    If UCase(strDocType) = UCase("Info Needed Letter") Then
                        pLicen.pLicenseeCourseTest.GetAll(Integer.Parse(Selrow.Cells("Licenseeid").Value))
                        Dim drRow As DataRow
                        Dim bolClosure As Boolean = True
                        Dim bolInstall As Boolean = True
                        Dim dtCourses As DataTable

                        'If Date.Compare(pLicen.ISSUED_DATE, CDate("01/01/0001")) = 0 Then
                        If Date.Compare(CDate(Selrow.Cells("Issued_date").Text), CDate("01/01/001")) = 0 Then
                            bolClosure = False
                            bolInstall = False
                        Else
                            dtCourses = pLicen.pLicenseeCourse.CourseTable
                            For Each drRow In dtCourses.Rows
                                If Date.Compare(drRow.Item("Date"), pLicen.ISSUED_DATE) > 0 Then
                                    If drRow.Item("CourseID") = 921 Then ' closure
                                        bolClosure = False
                                    ElseIf drRow.Item("CourseID") = 920 Then ' install
                                        bolInstall = False
                                    End If
                                End If
                            Next
                        End If

                        'Dim dtTests As DataTable = pLicen.pLicenseeCourseTest.TestTable
                        'For Each drRow In dtTests.Rows
                        '    If drRow.Item("Type") = 921 And drRow.Item("Score") < 75 Then
                        '        bolClosure = True
                        '    ElseIf drRow.Item("Type") = 920 And drRow.Item("Score") > 75 Then
                        '        bolInstall = True
                        '    End If
                        'Next
                        If bolClosure Then
                            If Selrow.Cells("LicenseeNo").Text.StartsWith("CHB") Or _
                                Selrow.Cells("LicenseeNo").Text.StartsWith("CRX") Or _
                                Selrow.Cells("LicenseeNo").Text.StartsWith("CHX") Or _
                                Selrow.Cells("LicenseeNo").Text.StartsWith("NHB") Or _
                                Selrow.Cells("LicenseeNo").Text.StartsWith("NRX") Or _
                                Selrow.Cells("LicenseeNo").Text.StartsWith("NHX") Then
                                count += 1
                                '    strInfoNeeded += count.ToString + ". A Closure course completion certificate" + vbCrLf
                            End If
                        End If

                        If bolInstall Then
                            If Selrow.Cells("LicenseeNo").Text.StartsWith("NHB") Or _
                                Selrow.Cells("LicenseeNo").Text.StartsWith("NRX") Or _
                                Selrow.Cells("LicenseeNo").Text.StartsWith("NHX") Then
                                count += 1
                                '      strInfoNeeded += count.ToString + ". An Install course completion certificate" + vbCrLf
                            End If
                        End If

                        If Selrow.Cells("LicenseeNo").Text.StartsWith("CHB") Or _
                            Selrow.Cells("LicenseeNo").Text.StartsWith("CRX") Or _
                            Selrow.Cells("LicenseeNo").Text.StartsWith("NHB") Or _
                            Selrow.Cells("LicenseeNo").Text.StartsWith("NRX") Then
                            If Not Selrow.Cells("hireStatus").Value Is System.DBNull.Value Then
                                If Selrow.Cells("hireStatus").Value.ToString.IndexOf("For Hire - Employee") > 0 Then
                                    If Selrow.Cells("Employee_Letter").Value = False Then
                                        count += 1
                                        '         strInfoNeeded += count.ToString + ". An Employee letter stating you are a full time employee of company" + vbCrLf
                                    End If
                                End If
                            End If
                        End If

                        If Selrow.Cells("LicenseeNo").Text.StartsWith("CHB") Or _
                            Selrow.Cells("LicenseeNo").Text.StartsWith("CHX") Or _
                            Selrow.Cells("LicenseeNo").Text.StartsWith("NHB") Or _
                            Selrow.Cells("LicenseeNo").Text.StartsWith("NHX") Then
                            If Not (Selrow.Cells("fin_resp_end_date").Text = String.Empty) Then
                                If Not Selrow.Cells("Expire_Date").Text = String.Empty Then
                                    If CDate(Selrow.Cells("fin_resp_end_date").Text) < CDate(Selrow.Cells("Expire_Date").Text) Then
                                        count += 1
                                        '   strInfoNeeded += count.ToString + ". Also if you work for hire(you have an ''H'' in your certification number), you must comply with the financial responsibility requirements in one of the following ways: " + vbCrLf
                                        '  strInfoNeeded += "      (1) submit to MDEQ a copy of an insurance certificate indicating that your company has atleast $50,000 of contractor's general liability insurance. This insurance certificate must list MDEQ as the certificate"
                                        ' strInfoNeeded += "holder and have 30- or 60-days cancellation notice." + vbCrLf + "           OR" + vbCrLf
                                        'strInfoNeeded += "       (2) submit to MDEQ a copy of your company's certificate of responsibility from the Mississippi Board of Contractors."

                                    End If
                                End If
                            End If
                        End If
                    End If

                    colParams.Add("<Licensee Greeting>", Selrow.Cells("Licensee").Value.ToString)
                    colParams.Add("<User>", strUserId)
                Else
                    colParams = Params
                End If


                Dim oWord As Word.Application = MusterContainer.GetWordApp

                If Not oWord Is Nothing Then

                    Dim strTempPath As String = TmpltPath & "Company\" & strTemplateName
                    Dim ltrGen As New MUSTER.BusinessLogic.pLetterGen
                    If strTempPath = TmpltPath & "Company\" & "InfoNeededLetter.doc" Then
                        If strInfoNeeded <> String.Empty Then
                            colParams.Add("<InfoNeeded>", strInfoNeeded)
                        Else
                            colParams.Add("<InfoNeeded>", String.Empty)
                        End If
                        ltrGen.CreateCompanyInfoLetter("Company", strDocName, colParams, strTempPath, Doc_Path & strDOC_NAME, oWord, strFile, strSignature)
                        oWord.Visible = True
                        UIUtilsGen.SaveDocument(nEntityID, UIUtilsGen.EntityTypes.Licensee, strDOC_NAME, strDocType, Doc_Path, strDocDesc, nModuleID, 0, 0, 0)
                    Else
                        ltrGen.CreateLetter("Company", strDocName, colParams, strTempPath, Doc_Path & strDOC_NAME, oWord, strFile, strSignature)
                        oWord.Visible = True
                        UIUtilsGen.SaveDocument(nEntityID, UIUtilsGen.EntityTypes.Licensee, strDOC_NAME, strDocType, Doc_Path, strDocDesc, nModuleID, 0, 0, 0)
                    End If
                End If
                oWord = Nothing

            Catch ex As Exception
                Throw ex
            End Try

        Catch ex As Exception
            Throw ex
        Finally
            SetCursorType(System.Windows.Forms.Cursors.Default)
        End Try
    End Function
    'Friend Function GenerateLicenseeLetter(ByVal EntityID As Integer, ByVal strDocType As String, ByVal strDocName As String, ByVal strDocDesc As String, ByVal strTemplateName As String, ByVal strUserId As String, Optional ByVal strPMHead As String = "", Optional ByVal strClosureHead As String = "", Optional ByVal Selrow As Infragistics.Win.UltraWinGrid.UltraGridRow = Nothing, Optional ByVal nCertifiedMailNo As Integer = 0, Optional ByVal Parameters As Specialized.NameValueCollection = Nothing, Optional ByVal strUserId As String = "") As Boolean

    '    Dim strDOC_NAME As String = String.Empty
    '    Dim strToday As String = String.Empty
    '    Dim nFacId As Integer = 0
    '    Dim oFacInfo As MUSTER.Info.FacilityInfo
    '    Dim oOwnerInfo As MUSTER.Info.OwnerInfo
    '    Dim oPersonaInfo As MUSTER.Info.PersonaInfo
    '    Dim oAddressInfo As MUSTER.Info.AddressInfo
    '    Dim colParams As New Specialized.NameValueCollection

    '    Try

    '        If DOC_PATH = "\" Then
    '            Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
    '        End If

    '        strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm"))
    '        strDOC_NAME = "COM_" + strDocName.Trim.ToString + "_" + CStr(Trim(EntityID.ToString)) + "_" + strToday + ".doc"

    '        Dim ServerPath As String = TmpltPath + "\Company"
    '        Dim Localpath As String = "C:\Templates\Company".Trim
    '        CopyTemplatesFromServerToLocal(ServerPath, Localpath)

    '        If Parameters Is Nothing Then
    '            'Build NameValueCollection with Tags and Values.

    '            colParams.Add("<Date>", Format(Now, "MMMM dd, yyyy"))
    '            colParams.Add("<Expiration Date>", Selrow.Cells("Expire_Date").Value.ToShortDateString)
    '            colParams.Add("<Licensee Name>", Selrow.Cells("Licensee").Value)
    '            colParams.Add("<Company Name>", Selrow.Cells("CompanyName").Value)
    '            colParams.Add("<Company Address1>", Selrow.Cells("Address1").Value)

    '            If Selrow.Cells("Address2").Value = String.Empty Then
    '                colParams.Add("<Company Address2>", Selrow.Cells("City").Value & ", " & Selrow.Cells("State").Value.TrimEnd & " " & Selrow.Cells("Zip").Value)
    '                colParams.Add("<City/State/Zip>", "")
    '            Else
    '                colParams.Add("<Company Address2>", Selrow.Cells("Address2").Value)
    '                colParams.Add("<City/State/Zip>", Selrow.Cells("City").Value & ", " & Selrow.Cells("State").Value.TrimEnd & " " & Selrow.Cells("Zip").Value)
    '            End If

    '            colParams.Add("<Licensee Greeting>", "Dear " + Selrow.Cells("Licensee").Value.ToString + ":")
    '            colParams.Add("<User>", strUserId)
    '        Else
    '            colParams = Parameters
    '        End If



    '        Try

    '            Dim strTempPath As String = Localpath & "\" & strTemplateName
    '            Dim ltrGen As New MUSTER.BusinessLogic.pLetterGen
    '            ltrGen.CreateLetter("Company", strDocName, colParams, strTempPath, Doc_Path & strDOC_NAME, MusterContainer.GetWordApp)
    '            UIUtilsGen.SaveDocument(EntityID, 27, strDOC_NAME, strDocType, Doc_Path, strDocDesc)
    '        Catch ex As Exception
    '            Throw ex
    '        End Try

    '    Catch ex As Exception
    '        Throw ex
    '    Finally
    '        SetCursorType(System.Windows.Forms.Cursors.Default)
    '    End Try
    'End Function

#End Region
#Region "Inspection Letters"
    Friend Function GenerateInspectionAnnouncementLetters(ByVal ownerID As Integer, ByVal strDocType As String, ByVal strDocName As String, ByVal strDocDesc As String, ByVal strTemplateName As String, ByVal facsForOwner As DataTable, ByVal own As MUSTER.BusinessLogic.pOwner, ByRef ltrGen As MUSTER.BusinessLogic.pLetterGen, ByVal progressBarValueIncrement As Single) As Boolean
        Dim strDOC_NAME As String = String.Empty
        Dim strToday As String = String.Empty
        Dim oPersonaInfo As MUSTER.Info.PersonaInfo
        Dim oAddressInfo As MUSTER.Info.AddressInfo
        Dim colParams As New Specialized.NameValueCollection
        nModuleID = UIUtilsGen.ModuleID.Inspection

        Try
            If DOC_PATH = "\" Then
                Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
            End If

            strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm"))
            strDOC_NAME = "INS_" + strDocName.Trim.ToString + "_" + ownerID.ToString + "_" + strToday + ".doc"

            'Build NameValueCollection with Tags and Values.
            colParams.Add("<Title>", strDocType.ToString)
            colParams.Add("<DATE>", Format(Now, "MMMM d, yyyy"))

            If own.OrganizationID > 0 Then
                oPersonaInfo = own.Organization
                colParams.Add("<Owner Name>", own.BPersona.Company)
                colParams.Add("<Owner Greeting>", own.BPersona.Company.Trim)
            Else
                oPersonaInfo = own.Persona
                colParams.Add("<Owner Name>", own.BPersona.FirstName.Trim & IIf(own.BPersona.MiddleName.Length > 0, " ", "") & own.BPersona.MiddleName.Trim & " " & own.BPersona.LastName.Trim & " " & own.BPersona.Suffix)
                colParams.Add("<Owner Greeting>", own.BPersona.FirstName.Trim & IIf(own.BPersona.MiddleName.Length > 0, " ", "") & own.BPersona.MiddleName.Trim & " " & own.BPersona.LastName.Trim & " " & own.BPersona.Suffix)
            End If

            oAddressInfo = own.Address()
            colParams.Add("<Owner Address 1>", oAddressInfo.AddressLine1)
            If oAddressInfo.AddressLine2 = String.Empty Then
                colParams.Add("<Owner Address 2>", oAddressInfo.City & ", " & oAddressInfo.State.TrimEnd & " " & oAddressInfo.Zip)
                colParams.Add("<City/State/Zip>", "")
            Else
                colParams.Add("<Owner Address 2>", oAddressInfo.AddressLine2)
                colParams.Add("<City/State/Zip>", oAddressInfo.City & ", " & oAddressInfo.State.TrimEnd & " " & oAddressInfo.Zip)
            End If

            colParams.Add("<Facility>", IIf(facsForOwner.Rows.Count > 1, "facilities", "facility"))
            colParams.Add("<User>", MusterContainer.AppUser.Name)
            colParams.Add("<User Phone>", CType(MusterContainer.AppUser.PhoneNumber, String))

            Dim oWordApp As Word.Application = MusterContainer.GetWordApp

            If Not oWordApp Is Nothing Then

                Dim strTempPath As String = TmpltPath & "Inspection\" & strTemplateName
                Dim returnVal As String = String.Empty
                ltrGen.CreateInspAnnouncementLetters(ownerID, colParams, strTempPath, Doc_Path + strDOC_NAME, facsForOwner, own, progressBarValueIncrement, UIUtilsGen.ModuleID.Inspection, MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID, oWordApp)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Function
                End If




                UIUtilsGen.SaveDocument(ownerID, UIUtilsGen.EntityTypes.Owner, strDOC_NAME, strDocType, Doc_Path, strDocDesc, nModuleID, 0, 0, 0)

                ' Make word visible
                oWordApp.Visible = True
            End If
            oWordApp = Nothing


        Catch ex As Exception
            Throw ex
        Finally

            SetCursorType(System.Windows.Forms.Cursors.Default)
        End Try
    End Function
    Friend Function GenerateInspectionCheckList(ByVal facID As Integer, ByVal strDocType As String, ByVal strDocName As String, ByVal strDocDesc As String, ByVal strTemplateName As String, ByVal oInsp As MUSTER.BusinessLogic.pInspection, ByVal dtTank As DataSet, ByVal dtPipe As DataSet, ByVal dsTerm As DataSet, ByRef ltrGen As MUSTER.BusinessLogic.pLetterGen, ByVal progress As Integer, ByVal strFlags As String, Optional ByVal pFacility As BusinessLogic.pFacility = Nothing, Optional ByVal comments As String = "", Optional ByVal designatedOperator As String = "", Optional ByVal imgs As Collections.ArrayList = Nothing) As Word.Application
        Dim strDOC_NAME As String = String.Empty
        Dim strToday As String = String.Empty
        Dim oPersonaInfo As MUSTER.Info.PersonaInfo
        Dim oAddressInfo As MUSTER.Info.AddressInfo
        Dim colParams As New Specialized.NameValueCollection
        Dim str As String = String.Empty
        Dim ds As DataSet
        nModuleID = UIUtilsGen.ModuleID.Inspection
        Try
            If DOC_PATH = "\" Then
                Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
            End If

            strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm"))
            strDOC_NAME = "INS_" + strDocName.Trim.ToString + "_" + oInsp.CheckListMaster.Owner.Facility.ID.ToString + "_" + strToday + ".doc"

            'Build NameValueCollection with Tags and Values.
            colParams.Add("<Title>", strDocType.ToString)
            With oInsp.CheckListMaster.Owner.Facilities

                oAddressInfo = .FacilityAddress

                colParams.Add("<Facility Name>", .Name)
                colParams.Add("<Facility ID>", .ID.ToString)
                colParams.Add("<Facility Address>", .FacilityAddresses.AddressLine1)
                If .FacilityAddress.City.ToUpper <> .FacilityAddress.PhsycalTown.ToUpper Then

                    colParams.Add("<Facility City>", String.Format("{0}   ({1})", .FacilityAddress.City, .FacilityAddress.PhsycalTown))

                Else
                    colParams.Add("<Facility City>", .FacilityAddress.City.ToString)

                End If
                colParams.Add("<Facility County>", .FacilityAddress.County.ToString)
                colParams.Add("<LA-D>", IIf(.LatitudeDegree = -1, String.Empty, .LatitudeDegree.ToString))
                colParams.Add("<LA-M>", IIf(.LatitudeMinutes = -1, String.Empty, .LatitudeMinutes.ToString))
                colParams.Add("<LA-S>", IIf(.LatitudeSeconds = -1, String.Empty, .LatitudeSeconds.ToString))
                colParams.Add("<LO-D>", IIf(.LongitudeDegree = -1, String.Empty, .LongitudeDegree.ToString))
                colParams.Add("<LO-M>", IIf(.LongitudeMinutes = -1, String.Empty, .LongitudeMinutes.ToString))
                colParams.Add("<LO-S>", IIf(.LongitudeSeconds = -1, String.Empty, .LongitudeSeconds.ToString))

            End With

            If oInsp.CheckListMaster.Owner.PersonID = 0 And oInsp.CheckListMaster.Owner.OrganizationID <> 0 Then
                ' org
                colParams.Add("<UST Owner>", oInsp.CheckListMaster.Owner.Organization.Company)
            ElseIf oInsp.CheckListMaster.Owner.PersonID <> 0 And oInsp.CheckListMaster.Owner.OrganizationID = 0 Then
                ' person
                colParams.Add("<UST Owner>", oInsp.CheckListMaster.Owner.Persona.LastName + ", " + oInsp.CheckListMaster.Owner.Persona.FirstName)
            End If
            colParams.Add("<Owners Rep>", oInsp.OwnersRep)
            colParams.Add("<Owner Address>", oInsp.CheckListMaster.Owner.Addresses.AddressLine1)
            colParams.Add("<Owner City>", oInsp.CheckListMaster.Owner.Addresses.City)
            colParams.Add("<Owner State>", oInsp.CheckListMaster.Owner.Addresses.State)
            If oInsp.CheckListMaster.Owner.Address.City.ToUpper <> oInsp.CheckListMaster.Owner.Address.PhsycalTown.ToUpper Then
                colParams.Add("<Owner Zip>", String.Format("{0}   ({1})", oInsp.CheckListMaster.Owner.Addresses.Zip, oInsp.CheckListMaster.Owner.Address.PhsycalTown))
            Else
                colParams.Add("<Owner Zip>", oInsp.CheckListMaster.Owner.Addresses.Zip)
            End If
            colParams.Add("<Owner Phone>", oInsp.CheckListMaster.Owner.PhoneNumberOne)
            colParams.Add("<DesignatedOperator>", IIf(designatedOperator = "" OrElse designatedOperator = "OWNER", "Owner", designatedOperator))

            If oInsp.CheckListMaster.Owner.Facilities.CAPCandidate Then
                colParams.Add("<CAP LEVEL>", "YES")
                'colParams.Add("<CAP Y>", "X")
                'colParams.Add("<CAP N>", "-")
            Else
                colParams.Add("<CAP LEVEL>", "NO")
                'colParams.Add("<CAP Y>", "-")
                'colParams.Add("<CAP N>", "X")
            End If
            'colParams.Add("<CAP SIGN UP>", "'TODO'")


            Dim strFees As String = String.Empty
            ds = oInsp.CheckListMaster.Owner.RunSQLQuery("exec spGetOwnerFacilityPastDueList " + oInsp.OwnerID.ToString + "," + oInsp.FacilityID.ToString + "")
            If Not ds Is Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 AndAlso ds.Tables(0).Rows(0)(3) > 0 Then
                strFees = ds.Tables(0).Rows(0)(3)
                strFees = strFees.Split(".")(0)
                colParams.Add("<Facility Fee>", "$" + strFees)
            Else
                colParams.Add("<Facility Fee>", "$0.00 ")
            End If
            strFees = String.Empty
            ds = oInsp.CheckListMaster.Owner.RunSQLQuery("exec spGetOwnerFacilityPastDueList " + oInsp.OwnerID.ToString + " ,NULL")

            If Not ds Is Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 AndAlso ds.Tables(0).Rows(0)(3) > 0 Then
                strFees = ds.Tables(0).Rows(0)(3)
                strFees = strFees.Split(".")(0)
                colParams.Add("<Owner Fee>", "$" + strFees)
            Else
                colParams.Add("<Owner Fee>", "$0.00 ")
            End If
            colParams.Add("<Flags>", strFlags)

            ds = oInsp.CheckListMaster.GetCLInspectionHistory
            If ds.Tables(0).Rows.Count > 0 Then
                Dim oUser As New MUSTER.BusinessLogic.pUser
                Dim dv As DataView = ds.Tables(0).DefaultView
                dv.Sort = "DATE INSPECTED DESC"
                If Not dv.Item(0)("DEQ INSPECTOR") Is DBNull.Value Then
                    oUser.Retrieve(dv.Item(0)("DEQ INSPECTOR"))
                    colParams.Add("<DEQ Inspector>", oUser.Name)
                Else
                    colParams.Add("<DEQ Inspector>", "____________________")
                End If
                'If ds.Tables(0).Rows(0)("TIME IN") Is DBNull.Value Then
                '    colParams.Add("<Time In>", "__________")
                'Else
                '    colParams.Add("<Time In>", ds.Tables(0).Rows(0)("TIME IN").ToString)
                'End If
                'If ds.Tables(0).Rows(0)("TIME OUT") Is DBNull.Value Then
                '    colParams.Add("<Time Out>", "__________")
                'Else
                '    colParams.Add("<Time Out>", ds.Tables(0).Rows(0)("TIME OUT").ToString)
                'End If
                If dv.Item(0)("DATE INSPECTED") Is DBNull.Value Then
                    colParams.Add("<Date Inspected>", "__________")
                Else
                    colParams.Add("<Date Inspected>", CType(dv.Item(0)("DATE INSPECTED"), Date).ToShortDateString)
                End If
            Else
                colParams.Add("<DEQ Inspector>", "____________________")
                'colParams.Add("<Time In>", "__________")
                'colParams.Add("<Time Out>", "__________")
                colParams.Add("<Date Inspected>", "__________")
            End If

            Dim strTempPath As String = TmpltPath & "Inspection\" & strTemplateName



            Dim longitude As Decimal = -1

            Dim latitude As Decimal = -1


            If Not pFacility Is Nothing Then

                With pFacility
                    If .LongitudeDegree > 0 AndAlso .LongitudeMinutes >= 0 AndAlso .LatitudeDegree > 0 AndAlso .LatitudeMinutes >= 0 AndAlso .LatitudeSeconds >= 0 AndAlso .LongitudeSeconds >= 0 Then
                        longitude = .LongitudeDegree + (.LongitudeMinutes / 60) + (IIf(.LongitudeSeconds >= 0, .LongitudeSeconds / 3600, 0))
                        latitude = .LatitudeDegree + (.LatitudeMinutes / 60) + (IIf(.LatitudeSeconds >= 0, .LatitudeSeconds / 3600, 0))
                    End If
                End With
            End If



            ' Dim MapAddress As MapAddress
            'Dim img1 As Image
            'Dim img2 As Image
            'Dim img3 As Image
            'Dim message As String = String.Empty

            'Try
            'If MsgBox("Would you like to add a map to the site to your inspection printout?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            'MapAddress = New MapAddress(oAddressInfo.AddressLine1, oAddressInfo.AddressLine2, oAddressInfo.PhsycalTown, oAddressInfo.State, longitude, latitude)
            'img1 = MapAddress.LoadImage(10)
            'img2 = MapAddress.LoadImage(13)
            'img3 = MapAddress.LoadImage(16)
            'End If

            'Catch ex As Exception
            '   Message = ex.Message

            'End Try

        Dim oWord As Word.Application = MusterContainer.GetWordApp

            If Not oWord Is Nothing Then


                oWord.Visible = False


                ltrGen.CreateInspCheckList(facID, colParams, strTempPath, Doc_Path + strDOC_NAME, oInsp, dtTank, dtPipe, dsTerm, progress, sketchpath, oWord, imgs, String.Empty, comments, designatedOperator)


                UIUtilsGen.SaveDocument(facID, UIUtilsGen.EntityTypes.Facility, strDOC_NAME, strDocType, Doc_Path, strDocDesc, nModuleID, 0, 0, 0)
            End If

            Return oWord

        Catch ex As Exception
            Throw ex

        End Try
    End Function
#End Region
#Region "CAE Letters"
    Private Sub BuildGeneralCAEOCEColParams(ByRef colParams As Specialized.NameValueCollection, ByVal ugOwnerRow As Infragistics.Win.UltraWinGrid.UltraGridRow, Optional ByVal manager As String = "", Optional ByVal phone As String = "")
        Dim oUserExecDirector As New MUSTER.BusinessLogic.pUser
        Try
            If colParams Is Nothing Then
                colParams = New Specialized.NameValueCollection
            End If

            colParams.Add("<Date>", Format(Now, "MMMM d, yyyy"))

            oUserExecDirector.RetrieveExecutiveDirector()
            colParams.Add("<Executive Director>", oUserExecDirector.Name)
            colParams.Add("<Executive Director Title case>", UIUtilsGen.TitleCaseString(oUserExecDirector.Name))
            colParams.Add("<Executive Director CAPS>", oUserExecDirector.Name.ToUpper)

            'Contact Name
            'Dim dtContacts As DataTable = GetXHAndXLContacts(ugOwnerRow.Cells("OWNER_ID").Value, uiutilsgen.EntityTypes.Owner, UIUtilsGen.ModuleID.CAE)
            'If dtContacts.Rows.Count > 0 Then
            '    Dim strContactName As String = dtContacts.Rows(0).Item("CONTACT_Name")
            '    If strContactName <> String.Empty Then
            '        colParams.Add("<Contact Name>", strContactName)
            '        colParams.Add("<Owner Greeting>", "Dear " + strContactName)
            '        bolownerGreetingsAdded = True
            '    Else
            '        colParams.Add("<Contact Name>", "")
            '    End If
            'Else
            '    colParams.Add("<Contact Name>", "")
            'End If
            colParams.Add("<Contact Name>", "<DeleteMe>")

            colParams.Add("<Owner Name>", ugOwnerRow.Cells("OWNERNAME").Text.Trim.ToUpper)
            If ugOwnerRow.Cells("ORGANIZATION_ID").Value Is DBNull.Value Then
                colParams.Add("<Owner Greeting>", ugOwnerRow.Cells("OWNERNAME").Text.Trim + ":")
            Else
                colParams.Add("<Owner Greeting>", ugOwnerRow.Cells("OWNERNAME").Text.Trim + ":")
            End If

            colParams.Add("<Owner Address 1>", ugOwnerRow.Cells("ADDRESS_LINE_ONE").Text.Trim.ToUpper)
            If ugOwnerRow.Cells("ADDRESS_TWO").Text.Trim = String.Empty Then
                colParams.Add("<Owner Address 2>", ugOwnerRow.Cells("CITY").Text.Trim.ToUpper + ", " + ugOwnerRow.Cells("STATE").Text.Trim.ToUpper + " " + ugOwnerRow.Cells("ZIP").Text.Trim.ToUpper)
                colParams.Add("<Owner City/State/Zip>", "<DeleteMe>")
            Else
                colParams.Add("<Owner Address 2>", ugOwnerRow.Cells("ADDRESS_TWO").Text.Trim.ToUpper)
                colParams.Add("<Owner City/State/Zip>", ugOwnerRow.Cells("CITY").Text.Trim.ToUpper + ", " + ugOwnerRow.Cells("STATE").Text.Trim.ToUpper + " " + ugOwnerRow.Cells("ZIP").Text.Trim.ToUpper)
            End If

            ' OCE Values
            Dim strPenaltyAmount As String = ""
            Dim strPolicyAmount As String = ""
            Dim dt As Date = CDate("01/01/0001")
            'Dim dt As Date = IIf(ugOwnerRow.Cells("OVERRIDE DUE DATE").Value Is DBNull.Value, IIf(ugOwnerRow.Cells("NEXT DUE DATE").Value Is DBNull.Value, CDate("01/01/0001"), ugOwnerRow.Cells("NEXT DUE DATE").Value), ugOwnerRow.Cells("OVERRIDE DUE DATE").Value)
            'If Date.Compare(dt, CDate("01/01/0001")) = 0 Then
            '    colParams.Add("<Due Date>", "")
            'Else
            '    colParams.Add("<Due Date>", dt.Date.ToShortDateString)
            'End If

            If Not ugOwnerRow.Cells("OVERRIDE AMOUNT").Value Is DBNull.Value Then
                If ugOwnerRow.Cells("OVERRIDE AMOUNT").Value < 0 Then
                    ugOwnerRow.Cells("OVERRIDE AMOUNT").Value = DBNull.Value
                End If
            End If
            If Not ugOwnerRow.Cells("SETTLEMENT AMOUNT").Value Is DBNull.Value Then
                If ugOwnerRow.Cells("SETTLEMENT AMOUNT").Value < 0 Then
                    ugOwnerRow.Cells("SETTLEMENT AMOUNT").Value = DBNull.Value
                End If
            End If
            If Not ugOwnerRow.Cells("POLICY AMOUNT").Value Is DBNull.Value Then
                If ugOwnerRow.Cells("POLICY AMOUNT").Value < 0 Then
                    ugOwnerRow.Cells("POLICY AMOUNT").Value = DBNull.Value
                End If
            End If

            If ugOwnerRow.Cells("OVERRIDE AMOUNT").Value Is DBNull.Value Then
                If ugOwnerRow.Cells("SETTLEMENT AMOUNT").Value Is DBNull.Value Then
                    If ugOwnerRow.Cells("POLICY AMOUNT").Value Is DBNull.Value Then
                        strPenaltyAmount = ""
                    ElseIf ugOwnerRow.Cells("POLICY AMOUNT").Value = -1 Then
                        strPenaltyAmount = ""
                    Else
                        strPenaltyAmount = "$" + ugOwnerRow.Cells("POLICY AMOUNT").Value.ToString.Split(".")(0)
                    End If
                ElseIf ugOwnerRow.Cells("SETTLEMENT AMOUNT").Value = -1 Then
                    If ugOwnerRow.Cells("POLICY AMOUNT").Value Is DBNull.Value Then
                        strPenaltyAmount = ""
                    ElseIf ugOwnerRow.Cells("POLICY AMOUNT").Value = -1 Then
                        strPenaltyAmount = ""
                    Else
                        strPenaltyAmount = "$" + ugOwnerRow.Cells("POLICY AMOUNT").Value.ToString.Split(".")(0)
                    End If
                Else
                    strPenaltyAmount = "$" + ugOwnerRow.Cells("SETTLEMENT AMOUNT").Value.ToString.Split(".")(0)
                End If
            ElseIf ugOwnerRow.Cells("OVERRIDE AMOUNT").Value = -1 Then
                If ugOwnerRow.Cells("SETTLEMENT AMOUNT").Value Is DBNull.Value Then
                    If ugOwnerRow.Cells("POLICY AMOUNT").Value Is DBNull.Value Then
                        strPenaltyAmount = ""
                    ElseIf ugOwnerRow.Cells("POLICY AMOUNT").Value = -1 Then
                        strPenaltyAmount = ""
                    Else
                        strPenaltyAmount = "$" + ugOwnerRow.Cells("POLICY AMOUNT").Value.ToString.Split(".")(0)
                    End If
                ElseIf ugOwnerRow.Cells("SETTLEMENT AMOUNT").Value = -1 Then
                    If ugOwnerRow.Cells("POLICY AMOUNT").Value Is DBNull.Value Then
                        strPenaltyAmount = ""
                    ElseIf ugOwnerRow.Cells("POLICY AMOUNT").Value = -1 Then
                        strPenaltyAmount = ""
                    Else
                        strPenaltyAmount = "$" + ugOwnerRow.Cells("POLICY AMOUNT").Value.ToString.Split(".")(0)
                    End If
                Else
                    strPenaltyAmount = "$" + ugOwnerRow.Cells("SETTLEMENT AMOUNT").Value.ToString.Split(".")(0)
                End If
            Else
                strPenaltyAmount = "$" + ugOwnerRow.Cells("OVERRIDE AMOUNT").Value.ToString.Split(".")(0)
            End If
            colParams.Add("<Penalty Amount>", strPenaltyAmount)

            If ugOwnerRow.Cells("POLICY AMOUNT").Value Is DBNull.Value Then
                strPolicyAmount = ""
            ElseIf ugOwnerRow.Cells("POLICY AMOUNT").Value = -1 Then
                strPolicyAmount = ""
            Else
                strPolicyAmount = "$" + ugOwnerRow.Cells("POLICY AMOUNT").Value.ToString.Split(".")(0)
            End If
            colParams.Add("<Policy Amount>", strPolicyAmount)

            dt = IIf(ugOwnerRow.Cells("WORKSHOP DATE").Value Is DBNull.Value, CDate("01/01/0001"), ugOwnerRow.Cells("WORKSHOP DATE").Value)
            If Date.Compare(dt, CDate("01/01/0001")) = 0 Then
                colParams.Add("<Workshop Date>", "")
            Else
                colParams.Add("<Workshop Date>", dt.Date.ToString("MMMM d, yyyy"))
            End If

            dt = IIf(ugOwnerRow.Cells("SHOW CAUSE HEARING DATE").Value Is DBNull.Value, CDate("01/01/0001"), ugOwnerRow.Cells("SHOW CAUSE HEARING DATE").Value)
            If Date.Compare(dt, CDate("01/01/0001")) = 0 Then
                colParams.Add("<Show Cause Hearing Date>", "")
            Else
                colParams.Add("<Show Cause Hearing Date>", dt.Date.ToString("MMMM d, yyyy"))
            End If

            dt = IIf(ugOwnerRow.Cells("COMMISSION HEARING DATE").Value Is DBNull.Value, CDate("01/01/0001"), ugOwnerRow.Cells("COMMISSION HEARING DATE").Value)
            If Date.Compare(dt, CDate("01/01/0001")) = 0 Then
                colParams.Add("<Commission Hearing Date>", "")
            Else
                colParams.Add("<Commission Hearing Date>", dt.Date.ToString("MMMM d, yyyy"))
            End If

            If ugOwnerRow.Cells("AGREED ORDER #").Value Is DBNull.Value Then
                colParams.Add("<Agreed Order Num>", "_______________")
            ElseIf ugOwnerRow.Cells("AGREED ORDER #").Text = String.Empty Then
                colParams.Add("<Agreed Order Num>", "_______________")
            Else
                colParams.Add("<Agreed Order Num>", ugOwnerRow.Cells("AGREED ORDER #").Text)
            End If

            If ugOwnerRow.Cells("ADMINISTRATIVE ORDER #").Value Is DBNull.Value Then
                colParams.Add("<Administrative Order Num>", "_______________")
            ElseIf ugOwnerRow.Cells("ADMINISTRATIVE ORDER #").Text = String.Empty Then
                colParams.Add("<Administrative Order Num>", "_______________")
            Else
                colParams.Add("<Administrative Order Num>", ugOwnerRow.Cells("ADMINISTRATIVE ORDER #").Text)
            End If

            colParams.Add("<Paid Amount>", IIf(ugOwnerRow.Cells("PAID AMOUNT").Value Is DBNull.Value, String.Empty, ugOwnerRow.Cells("PAID AMOUNT").Value.ToString))

            dt = IIf(ugOwnerRow.Cells("DATE RECEIVED").Value Is DBNull.Value, CDate("01/01/0001"), ugOwnerRow.Cells("DATE RECEIVED").Value)
            If Date.Compare(dt, CDate("01/01/0001")) = 0 Then
                colParams.Add("<Date Received>", "")
            Else
                colParams.Add("<Date Received>", dt.Date.ToString("MMMM d, yyyy"))
            End If

            colParams.Add("<Due Date>", DateAdd(DateInterval.Day, 90, Today).ToString("MMMM d, yyyy"))
            colParams.Add("<TodayPlus30Days>", DateAdd(DateInterval.Day, 30, Today).ToString("MMMM d, yyyy"))
            Dim userInfoLocal As MUSTER.Info.UserInfo
            userInfoLocal = MusterContainer.AppUser.RetrieveCAEHead()
            colParams.Add("<User Phone>", IIf(phone = String.Empty, CType(userInfoLocal.PhoneNumber, String), phone))
            colParams.Add("<User>", IIf(manager = String.Empty, userInfoLocal.ManagerName, manager))
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub BuildFacilities(ByRef colParams As Specialized.NameValueCollection, ByRef strFacsForAgreedOrder() As String, ByVal dtFacs As DataTable)
        Dim dtNull As Date = CDate("01/01/0001")
        Dim dt As Date = dtNull
        Dim dv As DataView
        Dim strFacility As String
        Try
            If Not dtFacs Is Nothing Then
                If dtFacs.Rows.Count > 0 Then
                    ReDim strFacsForAgreedOrder(dtFacs.Rows.Count - 1)
                    dv = dtFacs.DefaultView
                    dv.Sort = "FACILITY_ID"
                    For i As Integer = 0 To dtFacs.Rows.Count - 1
                        strFacility = String.Empty
                        strFacility += dv.Item(i)("FACILITY").ToString + " (I.D. #" + dv.Item(i)("FACILITY_ID").ToString + "), " + dv.Item(i)("ADDRESS_STATE_COUNTY")
                        If Date.Compare(dv.Item(i)("INSPECTEDON"), dtNull) <> 0 Then
                            If Date.Compare(dt, dtNull) = 0 Then
                                dt = dv.Item(i)("INSPECTEDON")
                            ElseIf Date.Compare(dv.Item(i)("INSPECTEDON"), dt) < 0 Then
                                dt = dv.Item(i)("INSPECTEDON")
                            End If
                        End If
                        strFacsForAgreedOrder(i) = strFacility
                    Next
                End If
            End If
            'If strFacility <> String.Empty Then
            '    strFacility = strFacility.Trim.TrimEnd(",")
            'End If
            'colParams.Add("<Facilities>", strFacility)
            If Date.Compare(dt, dtNull) = 0 Then
                colParams.Add("<InspectedOn>", "-N/A-")
            Else
                colParams.Add("<InspectedOn>", dt.ToShortDateString)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Friend Function GenerateCAEOCELetter(ByVal ugOwnerRow As Infragistics.Win.UltraWinGrid.UltraGridRow, ByVal dtFacs As DataTable, ByVal dtCitations As DataTable, ByVal dtDiscreps As DataTable, ByVal dtCorActions As DataTable, ByVal dtCOIFacs As DataTable, ByVal letterTemplatePropertyID As Integer, ByRef strOCEGeneratedLetterName As String, ByVal prevLetterDate As Date, ByVal returnVal As String, Optional ByVal pOCE As MUSTER.BusinessLogic.pOwnerComplianceEvent = Nothing) As Boolean
        Dim strDOC_NAME As String = String.Empty
        Dim strDocType As String = String.Empty
        Dim strDocDesc As String = String.Empty
        Dim strTemplateName As String = String.Empty
        Dim strTempPath As String = TmpltPath + "CAE\"
        Dim strCOCTempPath As String = TmpltPath + "CAE\COC.doc"
        'Dim strCOITempPath As String = TmpltPath + "CAE\COI.doc"
        Dim strToday As String = String.Empty
        Dim colParams As New Specialized.NameValueCollection
        Dim strFacsForAgreedOrder(0) As String

        Dim alFacsIndex As New ArrayList
        Dim alCitationsIndex As New ArrayList
        Dim alCorrectiveActionIndex As New ArrayList
        Dim alCorrectiveActionWithDueDateIndex As New ArrayList
        Dim alCorrectiveActionAddOnIndex As New ArrayList
        Dim alDiscrepsIndex As New ArrayList
        Dim alDiscrepsCorrectiveActionIndex As New ArrayList

        Dim bolCertifiedMailRequired As Boolean = False
        Dim bolCOCRequired As Boolean = False
        Dim bolHasAgreedOrder As Boolean = False
        Dim bolSaveLetterGenDate As Boolean = False
        Dim cl As CAP_Letters
        Dim doCapLetter As Boolean = False

        nModuleID = UIUtilsGen.ModuleID.CAE
        Try
            '  cl = New CAP_Letters

            If DOC_PATH = "\" Then
                Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
            End If

            strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm"))

            Dim manager As String = String.Empty
            Dim managerPhone As String = String.Empty

            If Not dtFacs Is Nothing AndAlso dtFacs.Rows.Count > 0 Then
                manager = MusterContainer.AppUser.Name
                managerPhone = MusterContainer.AppUser.PhoneNumber
            End If

            If managerPhone Is Nothing OrElse managerPhone.Length = 0 Then
                Dim us As New BusinessLogic.pUser
                us.RetrieveCAEHead()
                managerPhone = us.PhoneNumber.ToString
            End If


            BuildGeneralCAEOCEColParams(colParams, ugOwnerRow, manager, managerPhone)
            BuildFacilities(colParams, strFacsForAgreedOrder, dtFacs)

            If Date.Compare(prevLetterDate, CDate("01/01/0001")) = 0 Then
                If Date.Compare(ugOwnerRow.Cells("OCE DATE").Value, CDate("01/01/0001")) = 0 Then
                    colParams.Add("<Previous Letter Date>", "-N/A-")
                Else
                    colParams.Add("<Previous Letter Date>", CDate(ugOwnerRow.Cells("OCE DATE").Value).ToString("MMMM d, yyyy"))
                End If
            Else
                colParams.Add("<Previous Letter Date>", prevLetterDate.ToString("MMMM d, yyyy"))
            End If

            ' if multiple facs
            If dtFacs.Rows.Count > 1 Then
                colParams.Add("<system / systems>", "systems")
                colParams.Add("<system is / systems are>", "systems are")
                colParams.Add("<system was / systems were>", "system were")

                colParams.Add("<facility was / facilities were>", "facilities were")
                colParams.Add("<facility / facilities>", "facilities")
                colParams.Add("<facility has / facilities have>", "facilities have")
                colParams.Add("<facility is / facilities are>", "facilities are")
                colParams.Add("<an establishment / establishments>", "establishments")
                colParams.Add("<Certification / Certifications>", "Certifications")
                colParams.Add("<indicate / indicates>", "indicates")
                colParams.Add("<indicates / indicate>", "indicates")

                colParams.Add("<inspection / inspections>", "inspections")
                colParams.Add("<file / files>", "files")
            Else
                colParams.Add("<system / systems>", "system")
                colParams.Add("<system is / systems are>", "system is")
                colParams.Add("<system was / systems were>", "system was")

                colParams.Add("<facility was / facilities were>", "facility was")
                colParams.Add("<facility / facilities>", "facility")
                colParams.Add("<facility has / facilities have>", "facility has")
                colParams.Add("<facility is / facilities are>", "facility is")
                colParams.Add("<an establishment / establishments>", "an establishment")
                colParams.Add("<Certification / Certifications>", "Certification")
                colParams.Add("<indicate / indicates>", "indicate")
                colParams.Add("<indicates / indicate>", "indicate")

                colParams.Add("<inspection / inspections>", "inspection")
                colParams.Add("<file / files>", "file")
            End If

            ' if multiple violations
            If dtCitations.Rows.Count > 1 Then
                colParams.Add("<violation / violations>", "violations")
                colParams.Add("<citation / citations>", "citations")
                colParams.Add("<this citation / these citations>", "these citations")
                colParams.Add("<this citation is / these citations are>", "these citations are")

                colParams.Add("<citation has / citations have>", "citations have")
                colParams.Add("<action is / actions are>", "actions are")
                colParams.Add("<this is / these are>", "these are")
                colParams.Add("<has / have>", "have")
                colParams.Add("<is / are>", "are")
                colParams.Add("<was / were>", "were")
            Else
                colParams.Add("<violation / violations>", "violation")
                colParams.Add("<citation / citations>", "citation")
                colParams.Add("<this citation / these citations>", "this citation")
                colParams.Add("<citation has / citations have>", "citation has")
                colParams.Add("<this citation is / these citations are>", "this citation is")

                colParams.Add("<action is / actions are>", "action is")
                colParams.Add("<this is / these are>", "this is")
                colParams.Add("<has / have>", "has")
                colParams.Add("<is / are>", "is")
                colParams.Add("<was / were>", "was")
            End If

            ' if multiple discrepancies
            If dtDiscreps.Rows.Count > 1 Then
                colParams.Add("<discrepancy / discrepancies>", "discrepancies")
                colParams.Add("<discrepancy has / discrepancies have>", "discrepancies have")
            Else
                colParams.Add("<discrepancy / discrepancies>", "discrepancy")
                colParams.Add("<discrepancy has / discrepancies have>", "discrepancy has")
            End If

            colParams.Add("<year>", Today.Year.ToString)

            Select Case letterTemplatePropertyID

                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.DiscrepanciesOnly) ' 1
                    strDOC_NAME = "CAE_OCE_DiscrepanciesOnly" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCEDiscrepancy"
                    strDocDesc = "CAE OCE Discrepancy Letter"
                    strTemplateName = "C&E 1.doc"
                    strTempPath += strTemplateName

                    dtCitations = Nothing

                    alFacsIndex.Add(1)
                    alDiscrepsIndex.Add(2)
                    alDiscrepsCorrectiveActionIndex.Add(3)

                    bolCOCRequired = True
                    colParams.Add("<enclosed>", "Certification of Compliance")
                    bolCertifiedMailRequired = False
                    bolSaveLetterGenDate = True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.CAT3_NoPrior_NOV) ' 2
                    strDOC_NAME = "CAE_OCE_CAT3_NoPrior_NOV" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCECAT3_NoPrior_NOV"
                    strDocDesc = "CAE OCE CAT3 NoPrior NOV Letter"
                    strTemplateName = "C&E 2.doc"
                    strTempPath += strTemplateName

                    dtDiscreps = Nothing

                    alFacsIndex.Add(1)
                    alCitationsIndex.Add(2)
                    alCorrectiveActionIndex.Add(3)

                    bolCOCRequired = True
                    colParams.Add("<enclosed>", "Certification of Compliance")
                    bolCertifiedMailRequired = True
                    bolSaveLetterGenDate = True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.CAT2_NoPrior_NOV_Workshop) ' 3
                    strDOC_NAME = "CAE_OCE_CAT2_NoPrior_NOV_Workshop" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_CAT2_NoPrior_NOV_Workshop"
                    strDocDesc = "CAE OCE CAT2 NoPrior NOV Workshop Letter"
                    strTemplateName = "C&E 3.doc"
                    strTempPath += strTemplateName

                    dtDiscreps = Nothing

                    alFacsIndex.Add(1)
                    alCitationsIndex.Add(2)
                    alCorrectiveActionIndex.Add(3)

                    bolCOCRequired = True
                    colParams.Add("<enclosed>", "Certification of Compliance")
                    bolCertifiedMailRequired = True
                    bolSaveLetterGenDate = True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.CAT2_1_CAT3_NOV_Workshop, True) ' 3
                    strDOC_NAME = "CAE_OCE_CAT2_1_CAT3_NOV_Workshop" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_CAT2_1_CAT3_NOV_Workshop"
                    strDocDesc = "CAE OCE CAT2 1 CAT3 NOV Workshop Letter"
                    strTemplateName = "C&E 3.doc"
                    strTempPath += strTemplateName

                    dtDiscreps = Nothing

                    alFacsIndex.Add(1)
                    alCitationsIndex.Add(2)
                    alCorrectiveActionIndex.Add(3)

                    bolCOCRequired = True
                    colParams.Add("<enclosed>", "Certification of Compliance")
                    bolCertifiedMailRequired = True
                    bolSaveLetterGenDate = True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.CAT1_CAT1_CAT2_1_CAT3_NOV_AgreedOrder) ' 4
                    strDOC_NAME = "CAE_OCE_CAT1_CAT1_CAT2_1_CAT3_NOV_AgreedOrder" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_CAT1_CAT1_CAT2_1_CAT3_NOV_AgreedOrder"
                    strDocDesc = "CAE OCE CAT1 CAT1 CAT2 1 CAT3 NOV AgreedOrder Letter"
                    strTemplateName = "C&E 4.doc"
                    strTempPath += strTemplateName

                    dtDiscreps = Nothing

                    alFacsIndex.Add(1)
                    alCitationsIndex.Add(2)
                    alCorrectiveActionIndex.Add(3)
                    alCitationsIndex.Add(4)
                    alCorrectiveActionIndex.Add(5)
                    alCorrectiveActionAddOnIndex.Add(6)

                    bolCOCRequired = True
                    colParams.Add("<CommaEnclosed>", ", Certification of Compliance")
                    bolCertifiedMailRequired = True
                    bolHasAgreedOrder = True
                    bolSaveLetterGenDate = True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.CAT1_1_CAT3_NOV_Workshop_AgreedOrder) ' 5
                    strDOC_NAME = "CAE_OCE_CAT1_1_CAT3_NOV_Workshop_AgreedOrder" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_CAT1_1_CAT3_NOV_Workshop_AgreedOrder"
                    strDocDesc = "CAE OCE CAT1_1_CAT3_NOV_Workshop_AgreedOrder Letter"
                    strTemplateName = "C&E 5.doc"
                    strTempPath += strTemplateName

                    dtDiscreps = Nothing

                    alFacsIndex.Add(1)
                    alCitationsIndex.Add(2)
                    alCorrectiveActionIndex.Add(3)
                    alCorrectiveActionAddOnIndex.Add(4)
                    alCitationsIndex.Add(5)
                    alCorrectiveActionIndex.Add(6)
                    alCorrectiveActionAddOnIndex.Add(7)

                    bolCOCRequired = True
                    colParams.Add("<CommaEnclosed>", ", Certification of Compliance")
                    bolCertifiedMailRequired = True
                    bolHasAgreedOrder = True
                    bolSaveLetterGenDate = True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.CAT1_NoPrior_NOV_Workshop_AgreedOrder, True) ' 5
                    strDOC_NAME = "CAE_OCE_CAT1_NoPrior_NOV_Workshop_AgreedOrder" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_CAT1_NoPrior_NOV_Workshop_AgreedOrder"
                    strDocDesc = "CAE OCE CAT1_NoPrior_NOV_Workshop_AgreedOrder Letter"
                    strTemplateName = "C&E 5.doc"
                    strTempPath += strTemplateName

                    dtDiscreps = Nothing

                    alFacsIndex.Add(1)
                    alCitationsIndex.Add(2)
                    alCorrectiveActionIndex.Add(3)
                    alCorrectiveActionAddOnIndex.Add(4)
                    alCitationsIndex.Add(5)
                    alCorrectiveActionIndex.Add(6)
                    alCorrectiveActionAddOnIndex.Add(7)

                    bolCOCRequired = True
                    colParams.Add("<CommaEnclosed>", ", Certification of Compliance")
                    bolCertifiedMailRequired = True
                    bolHasAgreedOrder = True
                    bolSaveLetterGenDate = True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.CAT2_CAT1_CAT2_1_CAT3_NOV_AgreedOrder) ' 6
                    strDOC_NAME = "CAE_OCE_CAT1_NoPrior_NOV_Workshop_AgreedOrder" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_CAT1_NoPrior_NOV_Workshop_AgreedOrder"
                    strDocDesc = "CAE OCE CAT1_NoPrior_NOV_Workshop_AgreedOrder Letter"
                    strTemplateName = "C&E 6.doc"
                    strTempPath += strTemplateName

                    dtDiscreps = Nothing

                    alFacsIndex.Add(1)
                    alCitationsIndex.Add(2)
                    alCorrectiveActionIndex.Add(3)
                    alCorrectiveActionAddOnIndex.Add(4)
                    alCitationsIndex.Add(5)
                    alCorrectiveActionIndex.Add(6)
                    alCorrectiveActionAddOnIndex.Add(7)

                    bolCOCRequired = True
                    colParams.Add("<CommaEnclosed>", ", Certification of Compliance")
                    bolCertifiedMailRequired = True
                    bolHasAgreedOrder = True
                    bolSaveLetterGenDate = True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.CAT3_CAT1_CAT2_1_CAT3_NOV_AgreedOrder) ' 7
                    strDOC_NAME = "CAE_OCE_CAT1_NoPrior_NOV_Workshop_AgreedOrder" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_CAT1_NoPrior_NOV_Workshop_AgreedOrder"
                    strDocDesc = "CAE OCE CAT1_NoPrior_NOV_Workshop_AgreedOrder Letter"
                    strTemplateName = "C&E 7.doc"
                    strTempPath += strTemplateName

                    dtDiscreps = Nothing

                    alFacsIndex.Add(1)
                    alCitationsIndex.Add(2)
                    alCorrectiveActionIndex.Add(3)
                    alCorrectiveActionAddOnIndex.Add(4)
                    alCitationsIndex.Add(5)
                    alCorrectiveActionIndex.Add(6)
                    alCorrectiveActionAddOnIndex.Add(7)

                    bolCOCRequired = True
                    colParams.Add("<CommaEnclosed>", ", Certification of Compliance")
                    bolCertifiedMailRequired = True
                    bolHasAgreedOrder = True
                    bolSaveLetterGenDate = True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.CAT3_1_CAT3_NOV_Workshop) ' 8
                    strDOC_NAME = "CAE_OCE_CAT1_NoPrior_NOV_Workshop_AgreedOrder" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_CAT1_NoPrior_NOV_Workshop_AgreedOrder"
                    strDocDesc = "CAE OCE CAT1_NoPrior_NOV_Workshop_AgreedOrder Letter"
                    strTemplateName = "C&E 8.doc"
                    strTempPath += strTemplateName

                    dtDiscreps = Nothing

                    alFacsIndex.Add(1)
                    alCitationsIndex.Add(2)
                    alCorrectiveActionIndex.Add(3)

                    bolCOCRequired = True
                    colParams.Add("<CommaEnclosed>", ", Certification of Compliance")
                    bolCertifiedMailRequired = True
                    bolSaveLetterGenDate = True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.NOV_A_AgreedOrder) ' 9
                    strDOC_NAME = "CAE_OCE_NOV_A_AgreedOrder" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_NOV_A_AgreedOrder"
                    strDocDesc = "CAE OCE NOV_A_AgreedOrder Letter"
                    strTemplateName = "C&E 9.doc"
                    strTempPath += strTemplateName

                    dtDiscreps = Nothing
                    dtCOIFacs = Nothing
                    ' strCOITempPath = String.Empty

                    alFacsIndex.Add(1)
                    'alCitationsIndex.Add(2)
                    'alCorrectiveActionIndex.Add(3)
                    'alCorrectiveActionAddOnIndex.Add(4)

                    bolCertifiedMailRequired = True
                    bolHasAgreedOrder = True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.NOV_A_2ndNotice) ' 10
                    strDOC_NAME = "CAE_OCE_NOV_A_2ndNotice" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_NOV_A_2ndNotice"
                    strDocDesc = "CAE OCE NOV_A_2ndNotice Letter"
                    strTemplateName = "C&E 10.doc"
                    strTempPath += strTemplateName

                    dtCitations = Nothing
                    dtDiscreps = Nothing
                    dtCOIFacs = Nothing
                    'strCOITempPath = String.Empty

                    alFacsIndex.Add(1)

                    bolCertifiedMailRequired = True

                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.NOV_A_ShowCauseHearing) ' 11
                    strDOC_NAME = "CAE_OCE_NOV_A_ShowCauseHearing" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_NOV_A_ShowCauseHearing"
                    strDocDesc = "CAE OCE NOV_A_ShowCauseHearing Letter"
                    strTemplateName = "C&E 11.doc"
                    strTempPath += strTemplateName

                    dtCitations = Nothing
                    dtDiscreps = Nothing
                    dtCOIFacs = Nothing
                    ' strCOITempPath = String.Empty

                    alFacsIndex.Add(1)

                    bolCertifiedMailRequired = True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.NOV_AgreedOrder_B_2ndNotice) ' 12
                    strDOC_NAME = "CAE_OCE_NOV_AgreedOrder_B_2ndNotice" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_NOV_AgreedOrder_B_2ndNotice"
                    strDocDesc = "CAE OCE NOV_AgreedOrder_B_2ndNotice Letter"
                    strTemplateName = "C&E 12.doc"
                    strTempPath += strTemplateName

                    dtCitations = Nothing
                    dtDiscreps = Nothing
                    dtCOIFacs = Nothing
                    '  strCOITempPath = String.Empty

                    alFacsIndex.Add(1)

                    bolCertifiedMailRequired = True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.NOV_AgreedOrder_B_AgreedOrder) ' 13
                    strDOC_NAME = "CAE_OCE_NOV_AgreedOrder_B_AgreedOrder" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_NOV_AgreedOrder_B_AgreedOrder"
                    strDocDesc = "CAE OCE NOV_AgreedOrder_B_AgreedOrder Letter"
                    strTemplateName = "C&E 13.doc"
                    strTempPath += strTemplateName

                    dtDiscreps = Nothing
                    dtCOIFacs = Nothing
                    ' strCOITempPath = String.Empty

                    alFacsIndex.Add(1)
                    ' alCitationsIndex.Add(2)
                    ' alCorrectiveActionIndex.Add(3)
                    ' alCorrectiveActionAddOnIndex.Add(4)

                    bolCertifiedMailRequired = True
                    bolHasAgreedOrder = True


                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.NOV_AgreedOrder_AfterRedTag) ' 47
                    strDOC_NAME = "CAE_OCE_NOV_AO_AfterRedTag" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_NOV_AgreedOrder_AfterRedTag"
                    strDocDesc = "CAE OCE NOV_AgreedOrder_AfterRedTag Letter"
                    strTemplateName = "C&E 47.doc"
                    strTempPath += strTemplateName

                    dtDiscreps = Nothing
                    dtCOIFacs = Nothing
                    ' strCOITempPath = String.Empty
                    alFacsIndex.Add(1)
                    'alCitationsIndex.Add(2)
                    'alCorrectiveActionIndex.Add(3)
                    'alCorrectiveActionAddOnIndex.Add(4)

                    bolCertifiedMailRequired = True
                    bolHasAgreedOrder = True

                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.NOV_AgreedOrder_B_ShowCauseHearing) ' 14
                    strDOC_NAME = "CAE_OCE_NOV_AgreedOrder_B_ShowCauseHearing" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_NOV_AgreedOrder_B_ShowCauseHearing"
                    strDocDesc = "CAE OCE NOV_AgreedOrder_B_ShowCauseHearing Letter"
                    strTemplateName = "C&E 14.doc"
                    strTempPath += strTemplateName

                    dtCitations = Nothing
                    dtDiscreps = Nothing
                    dtCOIFacs = Nothing
                    ' strCOITempPath = String.Empty

                    alFacsIndex.Add(1)

                    bolCertifiedMailRequired = True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.NOV_Workshop_C_AgreedOrder_WhenCurrentStatusIsNew_WorkshopResultIsNoShowORChoseToPayAndEscalatedStatusIsAgreedOrder) ' 15
                    strDOC_NAME = "CAE_OCE_NOV_Workshop_C_AgreedOrder_WhenCurrentStatusIsNew_WorkshopResultIsNoShowORChoseToPayAndEscalatedStatusIsAgreedOrder" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_NOV_Workshop_C_AgreedOrder_WhenCurrentStatusIsNew_WorkshopResultIsNoShowORChoseToPayAndEscalatedStatusIsAgreedOrder"
                    strDocDesc = "CAE OCE NOV_Workshop_C_AgreedOrder_WhenCurrentStatusIsNew_WorkshopResultIsNoShowORChoseToPayAndEscalatedStatusIsAgreedOrder Letter"
                    strTemplateName = "C&E 15.doc"
                    strTempPath += strTemplateName

                    dtDiscreps = Nothing
                    dtCOIFacs = Nothing
                    ' strCOITempPath = String.Empty

                    alFacsIndex.Add(1)
                    ' alCitationsIndex.Add(2)
                    ' alCorrectiveActionIndex.Add(3)
                    ' alCorrectiveActionAddOnIndex.Add(4)

                    bolCertifiedMailRequired = True
                    bolHasAgreedOrder = True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.NOV_Workshop_C_AgreedOrder_WhenCurrentStatusIsNew_WorkshopResultIsPassORWaiveAndEscalatedStatusIsAgreedOrder) ' 16
                    strDOC_NAME = "CAE_OCE_NOV_Workshop_C_AgreedOrder_WhenCurrentStatusIsNew_WorkshopResultIsPassORWaiveAndEscalatedStatusIsAgreedOrder" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_NOV_Workshop_C_AgreedOrder_WhenCurrentStatusIsNew_WorkshopResultIsPassORWaiveAndEscalatedStatusIsAgreedOrder"
                    strDocDesc = "CAE OCE NOV_Workshop_C_AgreedOrder_WhenCurrentStatusIsNew_WorkshopResultIsPassORWaiveAndEscalatedStatusIsAgreedOrder Letter"
                    strTemplateName = "C&E 16.doc"
                    strTempPath += strTemplateName

                    dtDiscreps = Nothing
                    dtCOIFacs = Nothing
                    'strCOITempPath = String.Empty

                    alFacsIndex.Add(1)
                    '  alCitationsIndex.Add(2)
                    '  alCorrectiveActionIndex.Add(3)
                    '  alCorrectiveActionAddOnIndex.Add(4)

                    bolCertifiedMailRequired = True
                    bolHasAgreedOrder = True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.NOV_Workshop_C_AgreedOrder_WhenCurrentStatusIs2ndNoticeAndEscalatedStatusIsAgreedOrder) ' 17
                    strDOC_NAME = "CAE_OCE_NOV_Workshop_C_AgreedOrder_WhenCurrentStatusIs2ndNoticeAndEscalatedStatusIsAgreedOrder" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_NOV_Workshop_C_AgreedOrder_WhenCurrentStatusIs2ndNoticeAndEscalatedStatusIsAgreedOrder"
                    strDocDesc = "CAE OCE NOV_Workshop_C_AgreedOrder_WhenCurrentStatusIs2ndNoticeAndEscalatedStatusIsAgreedOrder Letter"
                    strTemplateName = "C&E 17.doc"
                    strTempPath += strTemplateName

                    dtDiscreps = Nothing
                    dtCOIFacs = Nothing
                    ' strCOITempPath = String.Empty

                    alFacsIndex.Add(1)
                    ' alCitationsIndex.Add(2)
                    ' alCorrectiveActionIndex.Add(3)
                    ' alCorrectiveActionAddOnIndex.Add(4)

                    bolCertifiedMailRequired = True
                    bolHasAgreedOrder = True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.NOV_Workshop_C_ShowCauseHearing) ' 18
                    strDOC_NAME = "CAE_OCE_NOV_Workshop_C_ShowCauseHearing" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_NOV_Workshop_C_ShowCauseHearing"
                    strDocDesc = "CAE OCE NOV_Workshop_C_ShowCauseHearing Letter"
                    strTemplateName = "C&E 18.doc"
                    strTempPath += strTemplateName

                    dtCitations = Nothing
                    dtDiscreps = Nothing
                    dtCOIFacs = Nothing
                    ' strCOITempPath = String.Empty

                    alFacsIndex.Add(1)

                    bolCertifiedMailRequired = True


                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.NOV_Workshop_C_2ndNotice) ' 19
                    strDOC_NAME = "CAE_OCE_NOV_Workshop_C_2ndNotice" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_NOV_Workshop_C_2ndNotice"
                    strDocDesc = "CAE OCE NOV_Workshop_C_2ndNotice Letter"
                    strTemplateName = "C&E 19.doc"
                    strTempPath += strTemplateName

                    dtCitations = Nothing
                    dtDiscreps = Nothing
                    dtCOIFacs = Nothing
                    'strCOITempPath = String.Empty

                    alFacsIndex.Add(1)

                    bolCertifiedMailRequired = True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.NOV_Workshop_AgreedOrder_D_AgreedOrder_WhenCurrentStatusIsNew_WorkshopResultIsNoShowORChoseToPayAndEscalatedStatusIsAgreedOrder) ' 20
                    strDOC_NAME = "CAE_OCE_NOV_Workshop_AgreedOrder_D_AgreedOrder_WhenCurrentStatusIsNew_WorkshopResultIsNoShowORChoseToPayAndEscalatedStatusIsAgreedOrder" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_NOV_Workshop_AgreedOrder_D_AgreedOrder_WhenCurrentStatusIsNew_WorkshopResultIsNoShowORChoseToPayAndEscalatedStatusIsAgreedOrder"
                    strDocDesc = "CAE OCE NOV_Workshop_AgreedOrder_D_AgreedOrder_WhenCurrentStatusIsNew_WorkshopResultIsNoShowORChoseToPayAndEscalatedStatusIsAgreedOrder Letter"
                    strTemplateName = "C&E 20.doc"
                    strTempPath += strTemplateName

                    dtDiscreps = Nothing
                    dtCOIFacs = Nothing
                    '  strCOITempPath = String.Empty

                    alFacsIndex.Add(1)
                    'alCitationsIndex.Add(2)
                    'alCorrectiveActionIndex.Add(3)
                    'alCorrectiveActionAddOnIndex.Add(4)

                    bolCertifiedMailRequired = True
                    bolHasAgreedOrder = True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.NOV_Workshop_AgreedOrder_D_AgreedOrder_WhenCurrentStatusIsNew_WorkshopResultIsPassORWaiveAndEscalatedStatusIsAgreedOrder) ' 21
                    strDOC_NAME = "CAE_OCE_NOV_Workshop_AgreedOrder_D_AgreedOrder_WhenCurrentStatusIsNew_WorkshopResultIsPassORWaiveAndEscalatedStatusIsAgreedOrder" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_NOV_Workshop_AgreedOrder_D_AgreedOrder_WhenCurrentStatusIsNew_WorkshopResultIsPassORWaiveAndEscalatedStatusIsAgreedOrder"
                    strDocDesc = "CAE OCE NOV_Workshop_AgreedOrder_D_AgreedOrder_WhenCurrentStatusIsNew_WorkshopResultIsPassORWaiveAndEscalatedStatusIsAgreedOrder Letter"
                    strTemplateName = "C&E 21.doc"
                    strTempPath += strTemplateName

                    dtDiscreps = Nothing
                    dtCOIFacs = Nothing
                    ' strCOITempPath = String.Empty

                    alFacsIndex.Add(1)
                    alCitationsIndex.Add(2)
                    alCorrectiveActionIndex.Add(3)
                    alCorrectiveActionAddOnIndex.Add(4)

                    bolCertifiedMailRequired = True
                    bolHasAgreedOrder = True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.StandAloneAgreedOrder) ' 48
                    strDOC_NAME = "CAE_OCE_ALONE_AGREED_ORDER" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_ALONE_AGREED_ORDER"
                    strDocDesc = "CAE OCE Stand Alone Agreed Order by Request Letter"
                    strTemplateName = "C&E 48.doc"
                    strTempPath += strTemplateName

                    dtDiscreps = Nothing
                    dtCOIFacs = Nothing
                    ' strCOITempPath = String.Empty

                    'alFacsIndex.Add(1)
                    alCitationsIndex.Add(1)
                    alCorrectiveActionIndex.Add(2)
                    alCorrectiveActionAddOnIndex.Add(3)

                    bolCertifiedMailRequired = True
                    bolHasAgreedOrder = True

                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.NOV_Workshop_AgreedOrder_D_2ndNotice) ' 22
                    strDOC_NAME = "CAE_OCE_NOV_Workshop_AgreedOrder_D_2ndNotice" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_NOV_Workshop_AgreedOrder_D_2ndNotice"
                    strDocDesc = "CAE OCE NOV_Workshop_AgreedOrder_D_2ndNotice Letter"
                    strTemplateName = "C&E 22.doc"
                    strTempPath += strTemplateName

                    dtCitations = Nothing
                    dtDiscreps = Nothing
                    dtCOIFacs = Nothing
                    ' strCOITempPath = String.Empty

                    alFacsIndex.Add(1)

                    bolCertifiedMailRequired = True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.NOV_Workshop_AgreedOrder_D_AgreedOrder_WhenCurrentStatusIs2ndNoticeAndEscalatedStatusIsAgreedOrder) ' 23
                    strDOC_NAME = "CAE_OCE_NOV_Workshop_AgreedOrder_D_AgreedOrder_WhenCurrentStatusIs2ndNoticeAndEscalatedStatusIsAgreedOrder" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_NOV_Workshop_AgreedOrder_D_AgreedOrder_WhenCurrentStatusIs2ndNoticeAndEscalatedStatusIsAgreedOrder"
                    strDocDesc = "CAE OCE NOV_Workshop_AgreedOrder_D_AgreedOrder_WhenCurrentStatusIs2ndNoticeAndEscalatedStatusIsAgreedOrder Letter"
                    strTemplateName = "C&E 23.doc"
                    strTempPath += strTemplateName

                    dtDiscreps = Nothing
                    dtCOIFacs = Nothing
                    ' strCOITempPath = String.Empty

                    alFacsIndex.Add(1)
                    alCitationsIndex.Add(2)
                    alCorrectiveActionIndex.Add(3)
                    alCorrectiveActionAddOnIndex.Add(4)

                    bolCertifiedMailRequired = True
                    bolHasAgreedOrder = True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.NOV_Workshop_AgreedOrder_D_ShowCauseHearing) ' 24
                    strDOC_NAME = "CAE_OCE_NOV_Workshop_AgreedOrder_D_ShowCauseHearing" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_NOV_Workshop_AgreedOrder_D_ShowCauseHearing"
                    strDocDesc = "CAE OCE NOV_Workshop_AgreedOrder_D_ShowCauseHearing Letter"
                    strTemplateName = "C&E 24.doc"
                    strTempPath += strTemplateName

                    dtCitations = Nothing
                    dtDiscreps = Nothing
                    dtCOIFacs = Nothing
                    ' strCOITempPath = String.Empty

                    alFacsIndex.Add(1)

                    bolCertifiedMailRequired = True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.Hearing_CommissionHearing_WhenCurrentStatusIsShowCauseHearing) ' 25
                    MsgBox("No template provided. Will continue as if letter was generated")
                    Return True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.Hearing_ShowCauseAgreedOrder) ' 26
                    strDOC_NAME = "CAE_OCE_Hearing_ShowCauseAgreedOrder" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_Hearing_ShowCauseAgreedOrder"
                    strDocDesc = "CAE OCE Hearing_ShowCauseAgreedOrder Letter"
                    strTemplateName = "C&E 26.doc"
                    strTempPath += strTemplateName

                    dtDiscreps = Nothing
                    dtCOIFacs = Nothing
                    ' strCOITempPath = String.Empty

                    alFacsIndex.Add(1)
                    alCitationsIndex.Add(2)
                    alCorrectiveActionIndex.Add(3)
                    alCorrectiveActionAddOnIndex.Add(4)

                    bolCertifiedMailRequired = True
                    bolHasAgreedOrder = True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.Hearing_CommissionHearing_WhenCurrentStatusIsShowCauseAgreedOrder) ' 27
                    MsgBox("No template provided. Will continue as if letter was generated")
                    Return True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.Hearing_CommissionHearingNFARescinded) ' 28
                    MsgBox("No template provided. Will continue as if letter was generated")
                    Return True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.Hearing_AdministrativeOrder) ' 29
                    MsgBox("No template provided. Will continue as if letter was generated")
                    Return True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.NFA_NFA) ' 30
                    Dim bolNFAOCE As Boolean = False
                    Dim bolNFAWorkshop As Boolean = False
                    Dim bolNFAOrder As Boolean = False

                    doCapLetter = True

                    dtCOIFacs = Nothing
                    ' strCOITempPath = String.Empty

                    ' 30c
                    If Not ugOwnerRow.Cells("AGREED ORDER #").Value Is DBNull.Value Then
                        If ugOwnerRow.Cells("AGREED ORDER #").Text <> String.Empty Then
                            bolNFAOrder = True
                            strTemplateName = "C&E 30c.doc"
                            bolHasAgreedOrder = True
                        End If
                        'ElseIf ugOwnerRow.Cells("AGREED ORDER #").Value > 0 Then
                        '    bolNFAOrder = True
                        '    strTemplateName = "C&E 30c.doc"
                    End If
                    If Not bolNFAOrder Then
                        ' 30b - if involves a workshop, workshop result = pass and paid amt is null
                        If ugOwnerRow.Cells("WORKSHOP_REQUIRED").Value Then
                            If Not ugOwnerRow.Cells("WORKSHOP RESULT").Value Is DBNull.Value Then
                                If ugOwnerRow.Cells("WORKSHOP RESULT").Value = 1011 Then ' pass
                                    If ugOwnerRow.Cells("PAID AMOUNT").Value Is DBNull.Value Then
                                        bolNFAWorkshop = True
                                        bolHasAgreedOrder = False
                                        strTemplateName = "C&E 30b.doc"
                                    ElseIf ugOwnerRow.Cells("PAID AMOUNT").Value <= 0 Then
                                        bolNFAWorkshop = True
                                        bolHasAgreedOrder = False
                                        strTemplateName = "C&E 30b.doc"
                                    End If
                                End If
                            End If
                        End If
                        If Not bolNFAWorkshop Then
                            ' 30a
                            If Not ugOwnerRow.Cells("WORKSHOP_REQUIRED").Value Then
                                bolNFAOCE = True
                                bolHasAgreedOrder = False
                                strTemplateName = "C&E 30a.doc"
                            End If
                            If Not bolNFAOCE Then
                                Dim bolAllCitationsDiscrep As Boolean = True
                                If Not ugOwnerRow.ChildBands Is Nothing Then
                                    If Not ugOwnerRow.ChildBands(0).Rows Is Nothing Then
                                        For Each ugChildRowLocal As Infragistics.Win.UltraWinGrid.UltraGridRow In ugOwnerRow.ChildBands(0).Rows ' facilities / citations
                                            If Not ugChildRowLocal.ChildBands Is Nothing Then
                                                If Not ugChildRowLocal.ChildBands(0).Rows Is Nothing Then
                                                    For Each ugGrandChildRowLocal As Infragistics.Win.UltraWinGrid.UltraGridRow In ugChildRowLocal.ChildBands(0).Rows
                                                        If ugGrandChildRowLocal.Cells("CITATION_ID").Value <> 19 Then ' Category Discrepancy
                                                            bolAllCitationsDiscrep = False
                                                            Exit For
                                                        End If
                                                    Next ' For Each ugGrandChildRowLocal 
                                                End If ' If Not ugChildRowLocal.ChildBands(0).Rows Is Nothing Then
                                            End If ' If Not ugChildRowLocal.ChildBands Is Nothing Then
                                            If Not bolAllCitationsDiscrep Then Exit For
                                        Next ' For Each ugChildRowLocal 
                                    End If ' If Not ugOwnerRow.ChildBands(0).Rows Is Nothing Then
                                End If ' If Not ugOwnerRow.ChildBands Is Nothing Then
                                If bolAllCitationsDiscrep Then
                                    bolNFAOCE = True
                                    bolHasAgreedOrder = False
                                    strTemplateName = "C&E 30a.doc"
                                End If

                                If Not bolNFAOCE Then
                                    If ugOwnerRow.Cells("OVERRIDE AMOUNT").Value Is DBNull.Value Then
                                        If ugOwnerRow.Cells("SETTLEMENT AMOUNT").Value Is DBNull.Value Then
                                            If ugOwnerRow.Cells("POLICY AMOUNT").Value Is DBNull.Value Then
                                                bolNFAOCE = True
                                                bolHasAgreedOrder = False
                                                strTemplateName = "C&E 30a.doc"
                                            ElseIf ugOwnerRow.Cells("POLICY AMOUNT").Value = -1 Then
                                                bolNFAOCE = True
                                                bolHasAgreedOrder = False
                                                strTemplateName = "C&E 30a.doc"
                                            End If
                                        ElseIf ugOwnerRow.Cells("SETTLEMENT AMOUNT").Value = -1 Then
                                            If ugOwnerRow.Cells("POLICY AMOUNT").Value Is DBNull.Value Then
                                                bolNFAOCE = True
                                                bolHasAgreedOrder = False
                                                strTemplateName = "C&E 30a.doc"
                                            ElseIf ugOwnerRow.Cells("POLICY AMOUNT").Value = -1 Then
                                                bolNFAOCE = True
                                                bolHasAgreedOrder = False
                                                strTemplateName = "C&E 30a.doc"
                                            End If
                                        End If
                                    ElseIf ugOwnerRow.Cells("OVERRIDE AMOUNT").Value = -1 Then
                                        If ugOwnerRow.Cells("SETTLEMENT AMOUNT").Value Is DBNull.Value Then
                                            If ugOwnerRow.Cells("POLICY AMOUNT").Value Is DBNull.Value Then
                                                bolNFAOCE = True
                                                bolHasAgreedOrder = False
                                                strTemplateName = "C&E 30a.doc"
                                            ElseIf ugOwnerRow.Cells("POLICY AMOUNT").Value = -1 Then
                                                bolNFAOCE = True
                                                bolHasAgreedOrder = False
                                                strTemplateName = "C&E 30a.doc"
                                            End If
                                        ElseIf ugOwnerRow.Cells("SETTLEMENT AMOUNT").Value = -1 Then
                                            If ugOwnerRow.Cells("POLICY AMOUNT").Value Is DBNull.Value Then
                                                bolNFAOCE = True
                                                bolHasAgreedOrder = False
                                                strTemplateName = "C&E 30a.doc"
                                            ElseIf ugOwnerRow.Cells("POLICY AMOUNT").Value = -1 Then
                                                bolNFAOCE = True
                                                bolHasAgreedOrder = False
                                                strTemplateName = "C&E 30a.doc"
                                            End If
                                        End If
                                    End If
                                End If ' If Not bolNFAOCE Then
                            End If ' If Not bolNFAOCE Then
                        End If
                    End If
                    If bolNFAOrder = False And bolNFAWorkshop = False And bolNFAOCE = False Then
                        bolNFAOCE = True
                        bolHasAgreedOrder = False
                        strTemplateName = "C&E 30a.doc"
                    End If
                    strDOC_NAME = "CAE_OCE_NFA" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_NFA"
                    strDocDesc = "CAE OCE NFA Letter"
                    strTempPath += strTemplateName
                    dtCitations = Nothing
                    dtDiscreps = Nothing
                    alFacsIndex.Add(1)
                    bolCertifiedMailRequired = False
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.NFARescind) ' 39
                    strDOC_NAME = "CAE_OCE_NFARescind" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_NFARescind"
                    strDocDesc = "CAE OCE NFARescind Letter"
                    strTemplateName = "C&E 39.doc"
                    strTempPath += strTemplateName
                    dtCitations = Nothing
                    dtDiscreps = Nothing
                    alFacsIndex.Add(1)
                    bolCertifiedMailRequired = False
                    doCapLetter = True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.ViolationWithin90Days_Discrepancy_WhenOwnerHasNonDiscrepCitation) ' 41
                    strDOC_NAME = "CAE_OCE_ViolationWithin90Days_Discrepancy_WhenOwnerHasNonDiscrepCitation" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCEViolationWithin90Days_Discrepancy_WhenOwnerHasNonDiscrepCitation"
                    strDocDesc = "CAE OCE Violation Within 90 Days Discrepancy When Owner Has Non Discrep Citation Letter"
                    strTemplateName = "C&E 41.doc"
                    strTempPath += strTemplateName

                    dtCitations = Nothing

                    alFacsIndex.Add(1)
                    alDiscrepsIndex.Add(2)
                    alDiscrepsCorrectiveActionIndex.Add(3)

                    bolCOCRequired = True
                    colParams.Add("<enclosed>", "Certification of Compliance")
                    bolCertifiedMailRequired = False
                    bolSaveLetterGenDate = True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.ViolationWithin90Days_Discrepancy_WhenOwnerHasDiscrepCitationsOnly) ' 42
                    strDOC_NAME = "CAE_OCE_ViolationWithin90Days_Discrepancy_WhenOwnerHasDiscrepCitationsOnly" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCEViolationWithin90Days_Discrepancy_WhenOwnerHasDiscrepCitationsOnly"
                    strDocDesc = "CAE OCE Violation Within 90 Days Discrepancy When Owner Has Discrep Citations Only Letter"
                    strTemplateName = "C&E 42.doc"
                    strTempPath += strTemplateName

                    dtCitations = Nothing

                    alFacsIndex.Add(1)
                    alDiscrepsIndex.Add(2)
                    alDiscrepsCorrectiveActionIndex.Add(3)

                    bolCOCRequired = True
                    colParams.Add("<enclosed>", "Certification of Compliance")
                    bolCertifiedMailRequired = False
                    bolSaveLetterGenDate = True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.NOV_AgreedOrder_AgreedOrder) ' 43
                    strDOC_NAME = "CAE_OCE_NOV_AgreedOrder_AgreedOrder" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_NOV_AgreedOrder_AgreedOrder"
                    strDocDesc = "CAE OCE NOV_AgreedOrder_AgreedOrder Letter"
                    strTemplateName = "C&E 43.doc"
                    strTempPath += strTemplateName

                    dtDiscreps = Nothing
                    dtCOIFacs = Nothing
                    ' strCOITempPath = String.Empty

                    alFacsIndex.Add(1)
                    alCitationsIndex.Add(2)
                    alCorrectiveActionIndex.Add(3)
                    alCorrectiveActionAddOnIndex.Add(4)

                    bolCertifiedMailRequired = True
                    bolHasAgreedOrder = True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.NOV_AgreedOrder_ShowCauseHearing) ' 43
                    strDOC_NAME = "CAE_OCE_NOV_AgreedOrder_ShowCauseHearing" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_NOV_AgreedOrder_ShowCauseHearing"
                    strDocDesc = "CAE OCE NOV_AgreedOrder_ShowCauseHearing Letter"
                    strTemplateName = "C&E 44.doc"
                    strTempPath += strTemplateName

                    dtCitations = Nothing
                    dtDiscreps = Nothing
                    dtCOIFacs = Nothing
                    ' strCOITempPath = String.Empty

                    alFacsIndex.Add(1)

                    bolCertifiedMailRequired = True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.RedTag_Notice) ' 45

                    strDOC_NAME = "CAE_OCE_RedTagNotification" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_redTagNotification"
                    strDocDesc = "CAE Red Tag Delivery Notice Letter"
                    strTemplateName = "C&E 45.doc"
                    strTempPath += strTemplateName
                    bolCOCRequired = True
                    strCOCTempPath = TmpltPath + "CAE\SELFCERT.doc"
                    dtDiscreps = Nothing

                    alFacsIndex.Add(1)
                    alCitationsIndex.Add(2)

                    bolCertifiedMailRequired = True
                    bolSaveLetterGenDate = True
                    dtCitations = Nothing
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.RedTag_Warning) ' 46

                    strDOC_NAME = "CAE_OCE_RedTagWarningNotification" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAEOCE_redTagWarningNotification"
                    strDocDesc = "CAE Red Tag Warning Letter"
                    strTemplateName = "C&E 46.doc"
                    strTempPath += strTemplateName

                    dtDiscreps = Nothing

                    alFacsIndex.Add(1)

                    bolCertifiedMailRequired = True
                    bolSaveLetterGenDate = True

                Case Else
                    MsgBox("Invalid OCE Letter Template Num")
                    Return False
            End Select

            'To Avoid duplicate creation of Letters
            If FileExists(DOC_PATH + strDOC_NAME) Then
                MsgBox(strDOC_NAME + "already exists")
                Return False
            Else
                Try
                    If bolCertifiedMailRequired Then
                        strCertifiedMail = String.Empty
                        certMail = New CertifiedMail
                        certMail.ShowDialog()
                        If strCertifiedMail = String.Empty Then
                            colParams.Add("<Certified Mail>", "")
                        Else
                            colParams.Add("<Certified Mail>", strCertifiedMail)
                        End If
                    End If
                    'UIUtilsGen.CreateAndSaveDocument("CAE", ugOwnerRow.Cells("OWNER_ID").Value, uiutilsgen.EntityTypes.Owner, DOC_PATH, strDOC_NAME, strDocType, strTempPath, Doc_Path, strDocDesc, colParams)
                    Dim ltrGen As New MUSTER.BusinessLogic.pLetterGen
                    Dim wordApp As Word.Application = MusterContainer.GetWordApp



                    ltrGen.GenerateCAELetter(colParams, IIf(bolHasAgreedOrder, strFacsForAgreedOrder, Nothing), strTempPath, strCOCTempPath, String.Empty, doc_path + strDOC_NAME, dtFacs, dtCitations, dtDiscreps, dtCorActions, dtCOIFacs, _
                                            alFacsIndex, alCitationsIndex, alCorrectiveActionIndex, alCorrectiveActionWithDueDateIndex, _
                                            alCorrectiveActionAddOnIndex, alDiscrepsIndex, alDiscrepsCorrectiveActionIndex, bolCOCRequired, wordApp)

                    UIUtilsGen.SaveDocument(ugOwnerRow.Cells("OWNER_ID").Value, UIUtilsGen.EntityTypes.Owner, strDOC_NAME, strDocType, Doc_Path, strDocDesc, nModuleID, ugOwnerRow.Cells("OCE_ID").Value, 0, UIUtilsGen.EntityTypes.CAEOwnerComplianceEvent)

                    If bolSaveLetterGenDate Then
                        ' letter generated date
                        ' only for the first created letter (OCE Creation / Modification), the date is saved / referred
                        If pOCE Is Nothing Then pOCE = New MUSTER.BusinessLogic.pOwnerComplianceEvent
                        pOCE.SaveLetterGeneratedDate(0, ugOwnerRow.Cells("OCE_ID").Value, UIUtilsGen.EntityTypes.CAEOwnerComplianceEvent, ugOwnerRow.Cells("PENDING_LETTER_TEMPLATE_NUM").Value, Today.Date, False, MusterContainer.pLetter.ID, UIUtilsGen.ModuleID.CAE, MusterContainer.AppUser.UserKey, returnVal)
                        If Not UIUtilsGen.HasRights(returnVal) Then
                            Return False
                        End If
                    End If

                    'documentID = MusterContainer.pLetter.ID
                    strOCEGeneratedLetterName = strDOC_NAME


                    ' If letterTemplatePropertyID = UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.NFA_NFA) Or _
                    ' letterTemplatePropertyID = UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.NFARescind) Then
                    '' Dim str As String = String.Empty
                    '' Dim fac As New DataAccess.FacilityDB

                    '' For Each facRow As DataRow In dtFacs.Rows
                    '' If fac.DBHasCIUTOSITanks(facRow("FACILITY_ID")) Then
                    '' Str = String.Format("{0}{1}{2}", Str, IIf(Str.Length = 0, "", ","), facRow("FACILITY_ID").ToString())
                    '' End If
                    '' Next

                    '' fac = Nothing

                    '' If Not Str() Is Nothing AndAlso Str.Length > 0 Then
                    '' cl.SetupSystemToGenerateCAPYearly(CAP_Letters.CapAnnualMode.CurrentSummary, False, Str, 0, Now.Year)
                    '' End If

                    '' End If


                    cl = Nothing

                    wordApp.Visible = True

                    Return True
                Catch ex As Exception
                    Throw ex
                End Try
            End If

        Catch ex As Exception
            Throw ex
            Return False
        Finally
            cl = Nothing
        End Try
    End Function
    Friend Function GenerateCAEFCENoDiscrepancyLetter(ByVal drOwner As DataRow, ByVal dtFacs As DataTable) As Boolean
        Dim strDOC_NAME As String = String.Empty
        Dim strDocType As String = String.Empty
        Dim strDocDesc As String = String.Empty
        Dim strTempPath As String = TmpltPath + "CAE\"
        Dim strToday As String = String.Empty
        Dim colParams As New Specialized.NameValueCollection
        Dim alFacsIndex As New ArrayList
        nModuleID = UIUtilsGen.ModuleID.CAE
        Try
            If DOC_PATH = "\" Then
                Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
            End If

            strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm"))
            strDOC_NAME = "CAE_FCE_NoDiscrepancy_" + drOwner("OWNER_ID").ToString + "_" + strToday + ".doc"

            'To Avoid duplicate creation of Letters
            If FileExists(DOC_PATH + strDOC_NAME) Then
                MsgBox(strDOC_NAME + "already exists")
                Return False
            Else
                'Build NameValueCollection with Tags and Values.

                Dim manager As String = String.Empty
                Dim managerPhone As String = String.Empty

                If Not dtFacs Is Nothing AndAlso dtFacs.Rows.Count > 0 Then
                    manager = MusterContainer.AppUser.Name
                    managerPhone = MusterContainer.AppUser.PhoneNumber
                End If

                If managerPhone Is Nothing OrElse managerPhone.Length = 0 Then
                    Dim us As New BusinessLogic.pUser
                    us.RetrieveCAEHead()
                    managerPhone = us.PhoneNumber.ToString
                End If

                Dim userInfoLocal As MUSTER.Info.UserInfo
                userInfoLocal = MusterContainer.AppUser.RetrieveCAEHead()
                colParams.Add("<User Phone>", IIf(managerPhone = String.Empty, CType(userInfoLocal.PhoneNumber, String), managerPhone))
                colParams.Add("<User>", IIf(manager = String.Empty, userInfoLocal.ManagerName, manager))



                ' if multiple facs
                If dtFacs.Rows.Count > 1 Then
                    colParams.Add("<system / systems>", "systems")
                    colParams.Add("<system is / systems are>", "systems are")
                    colParams.Add("<system was / systems were>", "system were")

                    colParams.Add("<facility was / facilities were>", "facilities were")
                    colParams.Add("<facility / facilities>", "facilities")
                    colParams.Add("<facility has / facilities have>", "facilities have")
                    colParams.Add("<facility is / facilities are>", "facilities are")
                    colParams.Add("<an establishment / establishments>", "establishments")
                    colParams.Add("<Certification / Certifications>", "Certifications")
                    colParams.Add("<indicate / indicates>", "indicates")
                    colParams.Add("<indicates / indicate>", "indicates")

                    colParams.Add("<inspection / inspections>", "inspections")
                    colParams.Add("<file / files>", "files")
                Else
                    colParams.Add("<system / systems>", "system")
                    colParams.Add("<system is / systems are>", "system is")
                    colParams.Add("<system was / systems were>", "system was")

                    colParams.Add("<facility was / facilities were>", "facility was")
                    colParams.Add("<facility / facilities>", "facility")
                    colParams.Add("<facility has / facilities have>", "facility has")
                    colParams.Add("<facility is / facilities are>", "facility is")
                    colParams.Add("<an establishment / establishments>", "an establishment")
                    colParams.Add("<Certification / Certifications>", "Certification")
                    colParams.Add("<indicate / indicates>", "indicate")
                    colParams.Add("<indicates / indicate>", "indicate")

                    colParams.Add("<inspection / inspections>", "inspection")
                    colParams.Add("<file / files>", "file")
                End If



                colParams.Add("<Title>", "Signature Needed Letter")
                colParams.Add("<Date>", Format(Now, "MMMM d, yyyy"))

                colParams.Add("<Owner Name>", drOwner("OWNERNAME").ToString.Trim)
                colParams.Add("<Owner Greeting>", drOwner("OWNERNAME").ToString.Trim + ":")

                colParams.Add("<Owner Address 1>", drOwner("ADDRESS_LINE_ONE").ToString.Trim)
                If drOwner("ADDRESS_TWO").ToString.Trim = String.Empty Then
                    colParams.Add("<Owner Address 2>", drOwner("CITY").ToString.Trim & ", " & drOwner("STATE").ToString.Trim & " " & drOwner("ZIP").ToString.Trim)
                    colParams.Add("<Owner City/State/Zip>", "<DeleteMe>")
                Else
                    colParams.Add("<Owner Address 2>", drOwner("ADDRESS_TWO").ToString.Trim)
                    colParams.Add("<Owner City/State/Zip>", drOwner("ADDRESS_TWO").ToString.Trim & ", " & drOwner("ADDRESS_TWO").ToString.Trim & " " & drOwner("ADDRESS_TWO").ToString.Trim)
                End If

                'Dim userInfoLocal2 As MUSTER.Info.UserInfo
                'userInfoLocal2 = MusterContainer.AppUser.RetrieveCAEHead()
                'colParams.Add("<User Phone>", CType(userInfoLocal2.PhoneNumber, String))
                'colParams.Add("<User>", userInfoLocal2.Name)

                'userInfoLocal2 = Nothing

                alFacsIndex.Add(1)
                Try
                    strDocType = "CAEFCENoDiscrepancy"
                    strDocDesc = "CAE FCE No Discrepancy Letter"
                    strTempPath = TmpltPath + "CAE\NoDiscrepancy.doc"
                    Dim ltrGen As New MUSTER.BusinessLogic.pLetterGen
                    Dim wordApp As Word.Application = MusterContainer.GetWordApp
                    'ltrGen.GenerateCAEFCENoDiscrepancyLetter(colParams, strTempPath, doc_path + strDOC_NAME, dtFacs, wordApp)
                    ltrGen.GenerateCAELetter(colParams, Nothing, strTempPath, String.Empty, String.Empty, doc_path + strDOC_NAME, dtFacs, Nothing, Nothing, Nothing, Nothing, alFacsIndex, New ArrayList, New ArrayList, New ArrayList, New ArrayList, New ArrayList, New ArrayList, False, wordApp)
                    UIUtilsGen.SaveDocument(drOwner("OWNER_ID"), UIUtilsGen.EntityTypes.Owner, strDOC_NAME, strDocType, Doc_Path, strDocDesc, nModuleID, 0, 0, 0)
                    wordApp.Visible = True
                    Return True
                Catch ex As Exception
                    Throw ex
                End Try
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Friend Sub GenerateCAENFARescindLetter(ByVal ugOwnerRow As Infragistics.Win.UltraWinGrid.UltraGridRow, ByVal dtFacs As DataTable, ByVal prevLetterDate As Date, ByVal bolMultipleCitations As Boolean)
        Dim strDOC_NAME As String = String.Empty
        Dim strDocType As String = String.Empty
        Dim strDocDesc As String = String.Empty
        Dim strTemplateName As String = String.Empty
        Dim strTempPath As String = TmpltPath + "CAE\"
        Dim strToday As String = String.Empty
        Dim colParams As New Specialized.NameValueCollection

        Dim alFacsIndex As New ArrayList
        Dim cl As CAP_Letters
        nModuleID = UIUtilsGen.ModuleID.CAE
        Try
            cl = New CAP_Letters
            If DOC_PATH = "\" Then
                Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
            End If

            strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm"))

            BuildGeneralCAEOCEColParams(colParams, ugOwnerRow)

            If Date.Compare(prevLetterDate, CDate("01/01/0001")) = 0 Then
                If Date.Compare(ugOwnerRow.Cells("OCE DATE").Value, CDate("01/01/0001")) = 0 Then
                    colParams.Add("<Previous Letter Date>", "-N/A-")
                Else
                    colParams.Add("<Previous Letter Date>", CDate(ugOwnerRow.Cells("OCE DATE").Value).ToString("MMMM d, yyyy"))
                End If
            Else
                colParams.Add("<Previous Letter Date>", prevLetterDate.ToString("MMMM d, yyyy"))
            End If

            ' if multiple facs
            If dtFacs.Rows.Count > 1 Then
                colParams.Add("<system / systems>", "systems")
                colParams.Add("<system is / systems are>", "systems are")

                colParams.Add("<facility was / facilities were>", "facilities were")
                colParams.Add("<facility / facilities>", "facilities")
                colParams.Add("<facility has / facilities have>", "facilities have")
                colParams.Add("<facility is / facilities are>", "facilities are")
                colParams.Add("<an establishment / establishments>", "establishments")
                colParams.Add("<Certification / Certifications>", "Certifications")
                colParams.Add("<indicate / indicates>", "indicates")
                colParams.Add("<inspection / inspections>", "inspections")
                colParams.Add("<file / files>", "files")
            Else
                colParams.Add("<system / systems>", "system")
                colParams.Add("<system is / systems are>", "system is")

                colParams.Add("<facility was / facilities were>", "facility was")
                colParams.Add("<facility / facilities>", "facility")
                colParams.Add("<facility has / facilities have>", "facility has")
                colParams.Add("<facility is / facilities are>", "facility is")
                colParams.Add("<an establishment / establishments>", "an establishment")
                colParams.Add("<Certification / Certifications>", "Certification")
                colParams.Add("<indicate / indicates>", "indicate")
                colParams.Add("<inspection / inspections>", "inspection")
                colParams.Add("<file / files>", "file")
            End If

            ' if multiple violations
            If bolMultipleCitations Then
                colParams.Add("<violation / violations>", "violations")
                colParams.Add("<citation / citations>", "citations")
                colParams.Add("<this citation / these citations>", "these citations")
                colParams.Add("<citation has / citations have>", "citations have")
                colParams.Add("<action is / actions are>", "actions are")
                colParams.Add("<this is / these are>", "these are")
                colParams.Add("<has / have>", "have")
                colParams.Add("<is / are>", "are")
                colParams.Add("<was / were>", "were")
            Else
                colParams.Add("<violation / violations>", "violation")
                colParams.Add("<citation / citations>", "citation")
                colParams.Add("<this citation / these citations>", "this citation")
                colParams.Add("<citation has / citations have>", "citation has")
                colParams.Add("<action is / actions are>", "action is")
                colParams.Add("<this is / these are>", "this is")
                colParams.Add("<has / have>", "has")
                colParams.Add("<is / are>", "is")
                colParams.Add("<was / were>", "was")
            End If

            colParams.Add("<year>", Today.Year.ToString)

            strDOC_NAME = "CAE_OCE_NFARescind" + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"
            strDocType = "CAEOCE_NFARescind"
            strDocDesc = "CAE OCE NFARescind Letter"
            strTemplateName = "C&E 39.doc"
            strTempPath += strTemplateName
            alFacsIndex.Add(1)

            'To Avoid duplicate creation of Letters
            If FileExists(DOC_PATH + strDOC_NAME) Then
                MsgBox(strDOC_NAME + "already exists")
                Exit Sub
            Else
                Dim ltrGen As New MUSTER.BusinessLogic.pLetterGen
                Dim wordApp As Word.Application = MusterContainer.GetWordApp

                ltrGen.GenerateCAELetter(colParams, Nothing, strTempPath, String.Empty, String.Empty, doc_path + strDOC_NAME, dtFacs, Nothing, Nothing, Nothing, Nothing, _
                    alFacsIndex, New ArrayList, New ArrayList, New ArrayList, _
                    New ArrayList, New ArrayList, New ArrayList, False, wordApp)


                Dim str As String = String.Empty
                Dim fac As New DataAccess.FacilityDB

                For Each facRow As DataRow In dtFacs.Rows
                    If fac.DBHasCIUTOSITanks(facRow("FACILITY_ID")) Then
                        str = String.Format("{0}{1}{2}", str, IIf(str.Length = 0, "", ","), facRow("FACILITY_ID").ToString())
                    End If
                Next

                fac = Nothing

                If Not str Is Nothing AndAlso str.Length > 0 Then
                    cl.SetupSystemToGenerateCAPYearly(CAP_Letters.CapAnnualMode.StaticByYear, False, str, 0, Now.Year)
                End If

                cl = Nothing


                UIUtilsGen.SaveDocument(ugOwnerRow.Cells("OWNER_ID").Value, UIUtilsGen.EntityTypes.Owner, strDOC_NAME, strDocType, Doc_Path, strDocDesc, nModuleID, ugOwnerRow.Cells("OCE_ID").Value, 0, UIUtilsGen.EntityTypes.CAEOwnerComplianceEvent)

                wordApp.Visible = True
                End If
        Catch ex As Exception
            Throw ex
        Finally
            cl = Nothing
        End Try
    End Sub

    Friend Function GenerateCAELCELetter(ByVal ugLicenseeRow As Infragistics.Win.UltraWinGrid.UltraGridRow, ByRef strErr As String, ByVal prevLetterDate As Date, ByRef documentID As Integer) As Boolean
        'Dim pLicensee As New MUSTER.BusinessLogic.pLicensee
        Dim pOwner As New MUSTER.BusinessLogic.pOwner
        Dim oFacInfo As MUSTER.Info.FacilityInfo

        Dim strDOC_NAME As String = String.Empty
        Dim strDocType As String = String.Empty
        Dim strDocDesc As String = String.Empty
        Dim strTemplateName As String = String.Empty
        Dim strTempPath As String = TmpltPath + "CAE\"
        Dim strToday As String = String.Empty
        Dim colParams As New Specialized.NameValueCollection
        Dim alFacsIndex As New ArrayList
        Dim alCitationsIndex As New ArrayList
        Dim dtCitations As New DataTable
        Dim dtFacs As New DataTable
        Dim dr, drFac As DataRow
        Dim bolCertifiedMailRequired As Boolean = False
        Dim bolHasAgreedOrder As Boolean = False

        nModuleID = UIUtilsGen.ModuleID.CAE

        Try
            If DOC_PATH = "\" Then
                Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
            End If

            dtCitations.Columns.Add("FACILITY_ID", GetType(Integer))
            dtCitations.Columns.Add("CITATION_INDEX", GetType(Integer))
            dtCitations.Columns.Add("CITATIONTEXT", GetType(String))
            dtCitations.Columns.Add("CorrectiveAction", GetType(String))
            dtCitations.Columns.Add("DUE", GetType(String))

            dtFacs.Columns.Add("FACILITY_ID", GetType(Integer))
            dtFacs.Columns.Add("FACILITY", GetType(String))
            dtFacs.Columns.Add("ADDRESS", GetType(String))

            strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm"))

            Dim i As Integer = 1
            For Each ugCitRow As Infragistics.Win.UltraWinGrid.UltraGridRow In ugLicenseeRow.ChildBands(0).Rows
                dr = dtCitations.NewRow
                dr("FACILITY_ID") = ugCitRow.Cells("FACILITY_ID").Value
                dr("CITATION_INDEX") = i
                i += 1
                dr("CITATIONTEXT") = ugCitRow.Cells("Citation Text").Text
                dr("CorrectiveAction") = String.Empty
                dr("DUE") = String.Empty
                dtCitations.Rows.Add(dr)
            Next


            Dim manager As String = String.Empty
            Dim managerPhone As String = String.Empty

            If Not dtFacs Is Nothing AndAlso dtFacs.Rows.Count > 0 Then
                manager = MusterContainer.AppUser.Name
                managerPhone = MusterContainer.AppUser.PhoneNumber
            End If

            If managerPhone Is Nothing OrElse managerPhone.Length = 0 Then
                Dim us As New BusinessLogic.pUser
                us.RetrieveCAEHead()
                managerPhone = us.PhoneNumber.ToString
            End If

            Dim userInfoLocal As MUSTER.Info.UserInfo
            userInfoLocal = MusterContainer.AppUser.RetrieveCAEHead()
            colParams.Add("<User Phone>", IIf(managerPhone = String.Empty, CType(userInfoLocal.PhoneNumber, String), managerPhone))
            colParams.Add("<User>", IIf(manager = String.Empty, userInfoLocal.ManagerName, manager))

            ' if multiple facs - there will always be only 1 fac
            colParams.Add("<system / systems>", "system")
            colParams.Add("<system is / systems are>", "system is")

            colParams.Add("<facility was / facilities were>", "facility was")
            colParams.Add("<facility / facilities>", "facility")
            colParams.Add("<facility has / facilities have>", "facility has")
            colParams.Add("<facility is / facilities are>", "facility is")
            colParams.Add("<an establishment / establishments>", "an establishment")
            colParams.Add("<Certification / Certifications>", "Certification")
            colParams.Add("<indicate / indicates>", "indicate")
            colParams.Add("<inspection / inspections>", "inspection")
            colParams.Add("<file / files>", "file")

            ' if multiple violations - there will always be only 1 violation
            colParams.Add("<violation / violations>", "violation")
            colParams.Add("<citation / citations>", "citation")
            colParams.Add("<this citation / these citations>", "this citation")
            colParams.Add("<citation has / citations have>", "citation has")
            colParams.Add("<action is / actions are>", "action is")
            colParams.Add("<this is / these are>", "this is")
            colParams.Add("<has / have>", "has")
            colParams.Add("<is / are>", "is")
            colParams.Add("<was / were>", "was")

            colParams.Add("<year>", Today.Year.ToString)

            Select Case ugLicenseeRow.Cells("PENDING_LETTER_TEMPLATE_NUM").Value
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.LCELetterTemplateNum.NOV)
                    strDOC_NAME = "CAE_LCE_NOV" + "_" + CStr(Trim(ugLicenseeRow.Cells("Licensee_id").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAELCE_NOV"
                    strDocDesc = "CAE LCE NOV Letter"
                    strTemplateName = "C&E 31.doc"
                    strTempPath += strTemplateName
                    alFacsIndex.Add(1)
                    alCitationsIndex.Add(2)
                    alCitationsIndex.Add(3)
                    bolCertifiedMailRequired = True
                    bolHasAgreedOrder = True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.LCELetterTemplateNum.NOV_ShowCauseHearing)
                    strDOC_NAME = "CAE_LCE_NOV_ShowCauseHearing" + "_" + CStr(Trim(ugLicenseeRow.Cells("Licensee_id").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAELCE_NOV_ShowCauseHearing"
                    strDocDesc = "CAE LCE NOV_ShowCauseHearing Letter"
                    strTemplateName = "C&E 32.doc"
                    strTempPath += strTemplateName
                    dtCitations = Nothing
                    alFacsIndex.Add(1)
                    bolCertifiedMailRequired = True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.LCELetterTemplateNum.Hearings_CommissionHearing_WhenCurrentStatusIsShowCauseHearingAndEscalatedStatusIsCommissionHearing)
                    MsgBox("No template provided. Will continue as if letter was generated")
                    Return True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.LCELetterTemplateNum.Hearings_ShowCauseAgreedOrder)
                    strDOC_NAME = "CAE_LCE_Hearings_ShowCauseAgreedOrder" + "_" + CStr(Trim(ugLicenseeRow.Cells("Licensee_id").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAELCE_Hearings_ShowCauseAgreedOrder"
                    strDocDesc = "CAE LCE Hearings_ShowCauseAgreedOrder Letter"
                    strTemplateName = "C&E 34.doc"
                    strTempPath += strTemplateName
                    alFacsIndex.Add(1)
                    alCitationsIndex.Add(2)
                    bolCertifiedMailRequired = True
                    bolHasAgreedOrder = True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.LCELetterTemplateNum.Hearings_CommissionHearing_WhenCurrentStatusIsShowCauseAgreedOrderAndEscalatedStatusIsCommissionHearing)
                    MsgBox("No template provided. Will continue as if letter was generated")
                    Return True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.LCELetterTemplateNum.Hearings_CommissionHearingNFARescinded)
                    MsgBox("No template provided. Will continue as if letter was generated")
                    Return True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.LCELetterTemplateNum.Hearings_AdministrativeOrder)
                    MsgBox("No template provided. Will continue as if letter was generated")
                    Return True
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.LCELetterTemplateNum.NFA_NFA)
                    strDOC_NAME = "CAE_LCE_NFA" + "_" + CStr(Trim(ugLicenseeRow.Cells("Licensee_id").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAELCE_NFA"
                    strDocDesc = "CAE LCE NFA Letter"
                    strTemplateName = "C&E 38.doc"
                    strTempPath += strTemplateName
                    dtCitations = Nothing
                    bolCertifiedMailRequired = False
                Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.LCELetterTemplateNum.NFARescind)
                    strDOC_NAME = "CAE_LCE_NFARescind" + "_" + CStr(Trim(ugLicenseeRow.Cells("Licensee_id").Value.ToString)) + "_" + strToday + ".doc"
                    strDocType = "CAELCE_NFARescind"
                    strDocDesc = "CAE LCE NFARescind Letter"
                    strTemplateName = "C&E 40.doc"
                    strTempPath += strTemplateName
                    dtCitations = Nothing
                    alFacsIndex.Add(1)
                    bolCertifiedMailRequired = False
                Case Else
                    strErr += ugLicenseeRow.Cells("Licensee").Text + " - Facility: " + ugLicenseeRow.Cells("Facility_ID").Text + vbCrLf
                    Return False
            End Select

            'To Avoid duplicate creation of Letters
            If FileExists(DOC_PATH + strDOC_NAME) Then
                MsgBox(strDOC_NAME + "already exists")
                Return False
            Else
                colParams.Add("<Title>", strDocType.ToString)
                colParams.Add("<Date>", Format(Now, "MMMM d, yyyy"))

                If Date.Compare(prevLetterDate, CDate("01/01/0001")) = 0 Then
                    If Date.Compare(ugLicenseeRow.Cells("LCE" + vbCrLf + "Date").Value, CDate("01/01/0001")) = 0 Then
                        colParams.Add("<Previous Letter Date>", "-N/A-")
                    Else
                        colParams.Add("<Previous Letter Date>", CDate(ugLicenseeRow.Cells("LCE" + vbCrLf + "Date").Value).ToShortDateString)
                    End If
                Else
                    colParams.Add("<Previous Letter Date>", prevLetterDate.ToShortDateString)
                End If

                Dim strPenaltyAmount As String = ""
                Dim strPolicyAmount As String = ""
                Dim dt As Date = CDate("01/01/0001")

                If Not ugLicenseeRow.Cells("Override" + vbCrLf + "Amount").Value Is DBNull.Value Then
                    If ugLicenseeRow.Cells("Override" + vbCrLf + "Amount").Value < 0 Then
                        ugLicenseeRow.Cells("Override" + vbCrLf + "Amount").Value = DBNull.Value
                    End If
                End If
                If Not ugLicenseeRow.Cells("Settlement" + vbCrLf + "Amount").Value Is DBNull.Value Then
                    If ugLicenseeRow.Cells("Settlement" + vbCrLf + "Amount").Value < 0 Then
                        ugLicenseeRow.Cells("Settlement" + vbCrLf + "Amount").Value = DBNull.Value
                    End If
                End If
                If Not ugLicenseeRow.Cells("Policy" + vbCrLf + "Amount").Value Is DBNull.Value Then
                    If ugLicenseeRow.Cells("Policy" + vbCrLf + "Amount").Value < 0 Then
                        ugLicenseeRow.Cells("Policy" + vbCrLf + "Amount").Value = DBNull.Value
                    End If
                End If

                If ugLicenseeRow.Cells("Override" + vbCrLf + "Amount").Value Is DBNull.Value Then
                    If ugLicenseeRow.Cells("Settlement" + vbCrLf + "Amount").Value Is DBNull.Value Then
                        If ugLicenseeRow.Cells("Policy" + vbCrLf + "Amount").Value Is DBNull.Value Then
                            strPenaltyAmount = ""
                        ElseIf ugLicenseeRow.Cells("Policy" + vbCrLf + "Amount").Value = -1 Then
                            strPenaltyAmount = ""
                        Else
                            strPenaltyAmount = "$" + ugLicenseeRow.Cells("Policy" + vbCrLf + "Amount").Value.ToString.Split(".")(0)
                        End If
                    ElseIf ugLicenseeRow.Cells("Settlement" + vbCrLf + "Amount").Value = -1 Then
                        If ugLicenseeRow.Cells("Policy" + vbCrLf + "Amount").Value Is DBNull.Value Then
                            strPenaltyAmount = ""
                        ElseIf ugLicenseeRow.Cells("Policy" + vbCrLf + "Amount").Value = -1 Then
                            strPenaltyAmount = ""
                        Else
                            strPenaltyAmount = "$" + ugLicenseeRow.Cells("Policy" + vbCrLf + "Amount").Value.ToString.Split(".")(0)
                        End If
                    Else
                        strPenaltyAmount = "$" + ugLicenseeRow.Cells("Settlement" + vbCrLf + "Amount").Value.ToString.Split(".")(0)
                    End If
                ElseIf ugLicenseeRow.Cells("Override" + vbCrLf + "Amount").Value = -1 Then
                    If ugLicenseeRow.Cells("Settlement" + vbCrLf + "Amount").Value Is DBNull.Value Then
                        If ugLicenseeRow.Cells("Policy" + vbCrLf + "Amount").Value Is DBNull.Value Then
                            strPenaltyAmount = ""
                        ElseIf ugLicenseeRow.Cells("Policy" + vbCrLf + "Amount").Value = -1 Then
                            strPenaltyAmount = ""
                        Else
                            strPenaltyAmount = "$" + ugLicenseeRow.Cells("Policy" + vbCrLf + "Amount").Value.ToString.Split(".")(0)
                        End If
                    ElseIf ugLicenseeRow.Cells("Settlement" + vbCrLf + "Amount").Value = -1 Then
                        If ugLicenseeRow.Cells("Policy" + vbCrLf + "Amount").Value Is DBNull.Value Then
                            strPenaltyAmount = ""
                        ElseIf ugLicenseeRow.Cells("Policy" + vbCrLf + "Amount").Value = -1 Then
                            strPenaltyAmount = ""
                        Else
                            strPenaltyAmount = "$" + ugLicenseeRow.Cells("Policy" + vbCrLf + "Amount").Value.ToString.Split(".")(0)
                        End If
                    Else
                        strPenaltyAmount = "$" + ugLicenseeRow.Cells("Settlement" + vbCrLf + "Amount").Value.ToString.Split(".")(0)
                    End If
                Else
                    strPenaltyAmount = "$" + ugLicenseeRow.Cells("Override" + vbCrLf + "Amount").Value.ToString.Split(".")(0)
                End If
                colParams.Add("<Penalty Amount>", strPenaltyAmount)

                If ugLicenseeRow.Cells("Policy" + vbCrLf + "Amount").Value Is DBNull.Value Then
                    strPolicyAmount = ""
                ElseIf ugLicenseeRow.Cells("Policy" + vbCrLf + "Amount").Value = -1 Then
                    strPolicyAmount = ""
                Else
                    strPolicyAmount = "$" + ugLicenseeRow.Cells("Policy" + vbCrLf + "Amount").Value.ToString.Split(".")(0)
                End If
                colParams.Add("<Policy Amount>", strPolicyAmount)

                colParams.Add("<Agreed Order Num>", "_______________")
                If ugLicenseeRow.Cells("Show Cause" + vbCrLf + "Hearing Date").Value Is DBNull.Value Then
                    colParams.Add("<Show Cause Hearing Date>", "")
                ElseIf Date.Compare(ugLicenseeRow.Cells("Show Cause" + vbCrLf + "Hearing Date").Value, CDate("01/01/0001")) = 0 Then
                    colParams.Add("<Show Cause Hearing Date>", "")
                Else
                    colParams.Add("<Show Cause Hearing Date>", CDate(ugLicenseeRow.Cells("Show Cause" + vbCrLf + "Hearing Date").Value).ToShortDateString)
                End If

                If ugLicenseeRow.Cells("Commission" + vbCrLf + "Hearing Date").Value Is DBNull.Value Then
                    colParams.Add("<Commission Hearing Date>", "")
                ElseIf Date.Compare(ugLicenseeRow.Cells("Commission" + vbCrLf + "Hearing Date").Value, CDate("01/01/0001")) = 0 Then
                    colParams.Add("<Commission Hearing Date>", "")
                Else
                    colParams.Add("<Commission Hearing Date>", CDate(ugLicenseeRow.Cells("Commission" + vbCrLf + "Hearing Date").Value).ToShortDateString)
                End If
                colParams.Add("<Due Date>", DateAdd(DateInterval.Day, 90, Today).ToShortDateString)
                colParams.Add("<TodayPlus30Days>", DateAdd(DateInterval.Day, 30, Today).ToShortDateString)

                Dim userInfoLocal2 As MUSTER.Info.UserInfo
                userInfoLocal2 = MusterContainer.AppUser.RetrieveCAEHead()
                colParams.Add("<User Phone>", CType(userInfoLocal2.PhoneNumber, String))
                colParams.Add("<User>", userInfoLocal2.Name)
                userInfoLocal2 = Nothing

                oFacInfo = pOwner.Facilities.Retrieve(pOwner.OwnerInfo, ugLicenseeRow.Cells("Facility_ID").Value, , "FACILITY")
                pOwner.Retrieve(oFacInfo.OwnerID)
                pOwner.OwnerInfo.facilityCollection = New MUSTER.Info.FacilityCollection
                pOwner.OwnerInfo.facilityCollection.Add(oFacInfo)
                ' to set the current instance to the facility id
                pOwner.Facilities.Retrieve(pOwner.OwnerInfo, oFacInfo.ID, , "FACILITY")

                colParams.Add("<Contact Name>", "")

                If pOwner.OrganizationID > 0 Then
                    colParams.Add("<Owner Name>", pOwner.BPersona.Company)
                    colParams.Add("<Owner Greeting>", pOwner.BPersona.Company.Trim + ":")
                Else
                    colParams.Add("<Owner Name>", pOwner.BPersona.Title.Trim & IIf(pOwner.BPersona.Title.Length > 0, " ", "") & pOwner.BPersona.FirstName.Trim & IIf(pOwner.BPersona.MiddleName.Trim.Length > 0, " ", "") & pOwner.BPersona.MiddleName.Trim & " " & pOwner.BPersona.LastName.Trim & " " & pOwner.BPersona.Suffix)
                    colParams.Add("<Owner Greeting>", pOwner.BPersona.FirstName.Trim & IIf(pOwner.BPersona.MiddleName.Trim.Length > 0, " ", "") & pOwner.BPersona.MiddleName.Trim & " " & pOwner.BPersona.LastName.Trim & " " & pOwner.BPersona.Suffix.Trim + ":")
                End If

                colParams.Add("<Owner Address 1>", pOwner.Address.AddressLine1)
                If pOwner.Address.AddressLine2 = String.Empty Then
                    colParams.Add("<Owner Address 2>", pOwner.Address.City & ", " & pOwner.Address.State.TrimEnd & " " & pOwner.Address.Zip)
                    colParams.Add("<Owner City/State/Zip>", "<DeleteMe>")
                Else
                    colParams.Add("<Owner Address 2>", pOwner.Address.AddressLine2)
                    colParams.Add("<Owner City/State/Zip>", pOwner.Address.City & ", " & pOwner.Address.State.TrimEnd & " " & pOwner.Address.Zip)
                End If

                dr = dtFacs.NewRow
                dr("FACILITY_ID") = pOwner.Facilities.ID
                dr("FACILITY") = pOwner.Facilities.Name
                dr("ADDRESS") = pOwner.Facilities.FacilityAddress.AddressLine1.Trim + ", " + _
                                pOwner.Facilities.FacilityAddress.City.Trim + " " + _
                                pOwner.Facilities.FacilityAddresses.State.Trim
                dtFacs.Rows.Add(dr)

                'colParams.Add("<Facility Name>", oFacInfo.Name)
                'colParams.Add("<Facility Address 1>", pOwner.Facilities.FacilityAddress.AddressLine1 & " " & IIf(pOwner.Facilities.FacilityAddress.AddressLine2 = String.Empty, "", pOwner.Facilities.FacilityAddress.AddressLine2))
                'If pOwner.Facilities.FacilityAddresses.AddressLine2 = String.Empty Then
                '    colParams.Add("<Facility Address 2>", pOwner.Facilities.FacilityAddress.City & ", " & pOwner.Facilities.FacilityAddress.State.TrimEnd & " " & pOwner.Facilities.FacilityAddress.Zip)
                '    colParams.Add("<City/State/Zip>", "")
                'Else
                '    colParams.Add("<Owner Address 2>", pOwner.Facilities.FacilityAddress.AddressLine2)
                '    colParams.Add("<City/State/Zip>", pOwner.Facilities.FacilityAddress.City & ", " & pOwner.Facilities.FacilityAddress.State.TrimEnd & " " & pOwner.Facilities.FacilityAddress.Zip)
                'End If

                Dim strFacsForAgreedOrder(0) As String
                strFacsForAgreedOrder(0) = pOwner.Facilities.Name + " (I.D. # " + pOwner.Facilities.ID.ToString + "), " + _
                                                pOwner.Facilities.FacilityAddress.AddressLine1 + ", " + _
                                                IIf(pOwner.Facilities.FacilityAddresses.AddressLine2 = String.Empty, "", pOwner.Facilities.FacilityAddresses.AddressLine2 + ", ") + _
                                                pOwner.Facilities.FacilityAddresses.City.Trim + " " + pOwner.Facilities.FacilityAddresses.State.Trim + " " + pOwner.Facilities.FacilityAddresses.Zip.Trim + " " + pOwner.Facilities.FacilityAddresses.County.Trim + " County"

                Try
                    If bolCertifiedMailRequired Then
                        strCertifiedMail = String.Empty
                        certMail = New CertifiedMail
                        certMail.ShowDialog()
                        If strCertifiedMail = String.Empty Then
                            colParams.Add("<Certified Mail>", "")
                        Else
                            colParams.Add("<Certified Mail>", strCertifiedMail)
                        End If
                    End If
                    Dim ltrGen As New MUSTER.BusinessLogic.pLetterGen
                    Dim wordApp As Word.Application = MusterContainer.GetWordApp
                    ltrGen.GenerateCAELetter(colParams, IIf(bolHasAgreedOrder, strFacsForAgreedOrder, Nothing), strTempPath, String.Empty, String.Empty, doc_path + strDOC_NAME, dtFacs, dtCitations, Nothing, Nothing, Nothing, alFacsIndex, alCitationsIndex, New ArrayList, New ArrayList, New ArrayList, New ArrayList, New ArrayList, False, wordApp)
                    UIUtilsGen.SaveDocument(ugLicenseeRow.Cells("Licensee_id").Value, UIUtilsGen.EntityTypes.CAELicenseeCompliantEvent, strDOC_NAME, strDocType, Doc_Path, strDocDesc, nModuleID, ugLicenseeRow.Cells("LCE_ID").Value, 0, UIUtilsGen.EntityTypes.CAELicenseeCompliantEvent)
                    documentID = MusterContainer.pLetter.ID
                    wordApp.Visible = True
                    Return True
                Catch ex As Exception
                    Throw ex
                End Try
            End If

        Catch ex As Exception
            Throw ex
            Return False
        End Try
    End Function
    'Friend Function GenerateCAELetter(ByVal facID As Integer, ByVal strDocType As String, ByVal strDocName As String, ByVal strDocDesc As String, ByVal strTemplateName As String) As Boolean
    '    'ByVal OwnerID As Integer, 
    '    Dim bolNeedCapDocs As Boolean = False
    '    Dim strDOC_NAME As String = String.Empty
    '    Dim strToday As String = String.Empty
    '    Dim nFacId As Integer = 0
    '    Dim oFacInfo As MUSTER.Info.FacilityInfo
    '    'Dim oOwnerInfo As MUSTER.Info.OwnerInfo
    '    Dim oPersonaInfo As MUSTER.Info.PersonaInfo
    '    Dim oAddressInfo As MUSTER.Info.AddressInfo
    '    Dim colParams As New Specialized.NameValueCollection
    '    Dim pOwner As New MUSTER.BusinessLogic.pOwner

    '    Try
    '        If DOC_PATH = "\" Then
    '            Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
    '        End If

    '        strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm"))
    '        strDOC_NAME = "CAE_" + strDocName.Trim.ToString + "_" + CStr(Trim(facID.ToString)) + "_" + strToday + ".doc"

    '        'Build NameValueCollection with Tags and Values.
    '        'oOwnerInfo = New MUSTER.Info.OwnerInfo
    '        oFacInfo = pOwner.Facilities.Retrieve(pOwner.OwnerInfo, facID, , "FACILITY")
    '        'oFacInfo = pOwner.Facilities.FacilityCollection.Item(facID)
    '        pOwner.Retrieve(oFacInfo.OwnerID)

    '        colParams.Add("<Title>", strDocType.ToString)
    '        colParams.Add("<DATE>", Format(Now, "MMMM dd, yyyy"))

    '        If pOwner.OrganizationID > 0 Then
    '            oPersonaInfo = pOwner.Organization
    '            colParams.Add("<Owner Name>", pOwner.BPersona.Company)
    '            colParams.Add("<Owner Greeting>", pOwner.BPersona.Company.Trim & ":")
    '        Else
    '            oPersonaInfo = pOwner.Persona
    '            colParams.Add("<Owner Name>", pOwner.BPersona.Title.Trim & IIf(pOwner.BPersona.Title.Length > 0, " ", "") & pOwner.BPersona.FirstName.Trim & IIf(pOwner.BPersona.MiddleName.Length > 0, " ", "") & pOwner.BPersona.MiddleName.Trim & " " & pOwner.BPersona.LastName.Trim & " " & pOwner.BPersona.Suffix)
    '            colParams.Add("<Owner Greeting>", "Dear " & pOwner.BPersona.Title.Trim & IIf(pOwner.BPersona.Title.Trim.Length > 0, " ", "") & pOwner.BPersona.FirstName.Trim & IIf(pOwner.BPersona.MiddleName.Trim.Length > 0, " ", "") & pOwner.BPersona.MiddleName.Trim & IIf(pOwner.BPersona.LastName.Trim.Length > 0, " ", "") & pOwner.BPersona.LastName.Trim & IIf(pOwner.BPersona.Suffix.Trim.Length > 0, " ", "") & pOwner.BPersona.Suffix.Trim & ":")
    '        End If

    '        oAddressInfo = pOwner.Address()
    '        colParams.Add("<Owner Address 1>", oAddressInfo.AddressLine1)
    '        If oAddressInfo.AddressLine2 = String.Empty Then
    '            colParams.Add("<Owner Address 2>", oAddressInfo.City & ", " & oAddressInfo.State.TrimEnd & " " & oAddressInfo.Zip)
    '            colParams.Add("<City/State/Zip>", "")
    '        Else
    '            colParams.Add("<Owner Address 2>", oAddressInfo.AddressLine2)
    '            colParams.Add("<City/State/Zip>", oAddressInfo.City & ", " & oAddressInfo.State.TrimEnd & " " & oAddressInfo.Zip)
    '        End If

    '        oAddressInfo = pOwner.Facilities.FacilityAddress
    '        colParams.Add("<Facility Name>", oFacInfo.Name)
    '        colParams.Add("<Facility Address>", oAddressInfo.AddressLine1 & " " & IIf(oAddressInfo.AddressLine2 = String.Empty, "", oAddressInfo.AddressLine2))
    '        colParams.Add("<Facility City>", oAddressInfo.City.TrimEnd)
    '        colParams.Add("<I.D. #>", oFacInfo.ID.ToString)
    '        'colParams.Add("<Due Date>", DueDate.ToShortDateString)
    '        'colParams.Add("<Schedule Date>", DueDate.ToShortDateString)

    '        colParams.Add("<User>", MusterContainer.AppUser.Name)

    '        Try
    '            Dim strTempPath As String = TmpltPath & "CAE\" & strTemplateName
    '            'Dim ltrGen As New MUSTER.BusinessLogic.pLetterGen
    '            'ltrGen.CreateInspAnnouncementLetters(ownerID, colParams, strTempPath, Doc_Path + strDOC_NAME, facsForOwner, own, progressBarValueIncrement)
    '            'SaveDocument(ownerID, 9, strDOC_NAME, strDocType, Doc_Path, strDocDesc)
    '            UIUtilsGen.CreateAndSaveDocument("CAE", facID, 6, DOC_PATH, strDOC_NAME, strDocType, strTempPath, Doc_Path, strDocDesc, colParams)
    '        Catch ex As Exception
    '            Throw ex
    '        End Try
    '    Catch ex As Exception
    '        Throw ex
    '    Finally
    '        SetCursorType(System.Windows.Forms.Cursors.Default)
    '    End Try
    'End Function
    'Friend Function GenerateOCELetter(ByVal ugOwnerRow As Infragistics.Win.UltraWinGrid.UltraGridRow, ByVal strDocType As String, ByVal strDocName As String, ByVal strDocDesc As String, ByVal strTemplateName As String, ByRef strOCEGeneratedLetterName As String) As Boolean
    '    Dim strDOC_NAME As String = String.Empty
    '    Dim strToday As String = String.Empty
    '    Dim colParams As New Specialized.NameValueCollection
    '    Dim pOwner As New MUSTER.BusinessLogic.pOwner

    '    Try
    '        If DOC_PATH = "\" Then
    '            Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
    '        End If

    '        strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm"))
    '        strDOC_NAME = strDocName.Trim.ToString + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"

    '        'Build NameValueCollection with Tags and Values.
    '        pOwner.Retrieve(ugOwnerRow.Cells("OWNER_ID").Value)

    '        colParams.Add("<Title>", strDocType.ToString)
    '        colParams.Add("<DATE>", Format(Now, "MMMM dd, yyyy"))

    '        colParams.Add("<Owner Name>", ugOwnerRow.Cells("OWNERNAME").Value.ToString)
    '        colParams.Add("<Owner Greeting>", ugOwnerRow.Cells("OWNERNAME").Value.ToString.Trim + ":")

    '        colParams.Add("<Owner Address 1>", pOwner.Address.AddressLine1)
    '        If pOwner.Address.AddressLine2 = String.Empty Then
    '            colParams.Add("<Owner Address 2>", pOwner.Address.City + ", " + pOwner.Address.State.TrimEnd + " " + pOwner.Address.Zip)
    '            colParams.Add("<City/State/Zip>", "")
    '        Else
    '            colParams.Add("<Owner Address 2>", pOwner.Address.AddressLine2)
    '            colParams.Add("<City/State/Zip>", pOwner.Address.City + ", " + pOwner.Address.State.TrimEnd + " " + pOwner.Address.Zip)
    '        End If

    '        ' OCE Values
    '        Dim strPenaltyAmount As String = ""
    '        Dim dt As Date = IIf(ugOwnerRow.Cells("OVERRIDE DUE DATE").Value Is DBNull.Value, IIf(ugOwnerRow.Cells("NEXT DUE DATE").Value Is DBNull.Value, CDate("01/01/0001"), ugOwnerRow.Cells("NEXT DUE DATE").Value), ugOwnerRow.Cells("OVERRIDE DUE DATE").Value)
    '        If Date.Compare(dt, CDate("01/01/0001")) = 0 Then
    '            colParams.Add("<Due Date>", "")
    '        Else
    '            colParams.Add("<Due Date>", dt.Date.ToShortDateString)
    '        End If

    '        If Not ugOwnerRow.Cells("OVERRIDE AMOUNT").Value Is DBNull.Value Then
    '            If ugOwnerRow.Cells("OVERRIDE AMOUNT").Value < 0 Then
    '                ugOwnerRow.Cells("OVERRIDE AMOUNT").Value = DBNull.Value
    '            End If
    '        End If
    '        If Not ugOwnerRow.Cells("SETTLEMENT AMOUNT").Value Is DBNull.Value Then
    '            If ugOwnerRow.Cells("SETTLEMENT AMOUNT").Value < 0 Then
    '                ugOwnerRow.Cells("SETTLEMENT AMOUNT").Value = DBNull.Value
    '            End If
    '        End If
    '        If Not ugOwnerRow.Cells("POLICY AMOUNT").Value Is DBNull.Value Then
    '            If ugOwnerRow.Cells("POLICY AMOUNT").Value < 0 Then
    '                ugOwnerRow.Cells("POLICY AMOUNT").Value = DBNull.Value
    '            End If
    '        End If

    '        If ugOwnerRow.Cells("OVERRIDE AMOUNT").Value Is DBNull.Value Then
    '            If ugOwnerRow.Cells("SETTLEMENT AMOUNT").Value Is DBNull.Value Then
    '                If ugOwnerRow.Cells("POLICY AMOUNT").Value Is DBNull.Value Then
    '                    strPenaltyAmount = ""
    '                Else
    '                    strPenaltyAmount = ugOwnerRow.Cells("POLICY AMOUNT").Value.ToString
    '                End If
    '            Else
    '                strPenaltyAmount = ugOwnerRow.Cells("SETTLEMENT AMOUNT").Value.ToString
    '            End If
    '        Else
    '            strPenaltyAmount = ugOwnerRow.Cells("OVERRIDE AMOUNT").Value.ToString
    '        End If
    '        colParams.Add("<Penalty Amount>", strPenaltyAmount)

    '        dt = IIf(ugOwnerRow.Cells("WORKSHOP DATE").Value Is DBNull.Value, CDate("01/01/0001"), ugOwnerRow.Cells("WORKSHOP DATE").Value)
    '        If Date.Compare(dt, CDate("01/01/0001")) = 0 Then
    '            colParams.Add("<Workshop Date>", "")
    '        Else
    '            colParams.Add("<Workshop Date>", dt.Date.ToShortDateString)
    '        End If

    '        dt = IIf(ugOwnerRow.Cells("SHOW CAUSE HEARING DATE").Value Is DBNull.Value, CDate("01/01/0001"), ugOwnerRow.Cells("SHOW CAUSE HEARING DATE").Value)
    '        If Date.Compare(dt, CDate("01/01/0001")) = 0 Then
    '            colParams.Add("<Show Cause Hearing Date>", "")
    '        Else
    '            colParams.Add("<Show Cause Hearing Date>", dt.Date.ToShortDateString)
    '        End If

    '        dt = IIf(ugOwnerRow.Cells("COMMISSION HEARING DATE").Value Is DBNull.Value, CDate("01/01/0001"), ugOwnerRow.Cells("COMMISSION HEARING DATE").Value)
    '        If Date.Compare(dt, CDate("01/01/0001")) = 0 Then
    '            colParams.Add("<Commission Hearing Date>", "")
    '        Else
    '            colParams.Add("<Commission Hearing Date>", dt.Date.ToShortDateString)
    '        End If

    '        colParams.Add("<Paid Amount>", IIf(ugOwnerRow.Cells("PAID AMOUNT").Value Is DBNull.Value, String.Empty, ugOwnerRow.Cells("PAID AMOUNT").Value.ToString))

    '        dt = IIf(ugOwnerRow.Cells("DATE RECEIVED").Value Is DBNull.Value, CDate("01/01/0001"), ugOwnerRow.Cells("DATE RECEIVED").Value)
    '        If Date.Compare(dt, CDate("01/01/0001")) = 0 Then
    '            colParams.Add("<Date Received>", "")
    '        Else
    '            colParams.Add("<Date Received>", dt.Date.ToShortDateString)
    '        End If

    '        colParams.Add("<User>", MusterContainer.AppUser.Name)

    '        Try
    '            Dim strTempPath As String = TmpltPath & "CAE\" & strTemplateName
    '            UIUtilsGen.CreateAndSaveDocument("CAE", ugOwnerRow.Cells("OWNER_ID").Value, uiutilsgen.EntityTypes.Owner, DOC_PATH, strDOC_NAME, strDocType, strTempPath, Doc_Path, strDocDesc, colParams)
    '            strOCEGeneratedLetterName = strDOC_NAME
    '        Catch ex As Exception
    '            Throw ex
    '        End Try
    '        ' need to return true if letter was generated successfully for further processing
    '        Return True
    '    Catch ex As Exception
    '        Throw ex
    '    Finally
    '        SetCursorType(System.Windows.Forms.Cursors.Default)
    '    End Try
    'End Function
    'Friend Function GenerateOCENFARescindLetter(ByVal ugOwnerRow As Infragistics.Win.UltraWinGrid.UltraGridRow, ByVal strDocType As String, ByVal strDocName As String, ByVal strDocDesc As String, ByVal strTemplateName As String, ByRef strOCEGeneratedLetterName As String) As Boolean
    '    Dim strDOC_NAME As String = String.Empty
    '    Dim strToday As String = String.Empty
    '    Dim colParams As New Specialized.NameValueCollection
    '    Dim pOwner As New MUSTER.BusinessLogic.pOwner

    '    Try
    '        If DOC_PATH = "\" Then
    '            Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
    '        End If

    '        strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm"))
    '        strDOC_NAME = strDocName.Trim.ToString + "_" + CStr(Trim(ugOwnerRow.Cells("OWNER_ID").Value.ToString)) + "_" + strToday + ".doc"

    '        'Build NameValueCollection with Tags and Values.
    '        pOwner.Retrieve(ugOwnerRow.Cells("OWNER_ID").Value)

    '        colParams.Add("<Title>", strDocType.ToString)
    '        colParams.Add("<DATE>", Format(Now, "MMMM dd, yyyy"))

    '        colParams.Add("<Owner Name>", ugOwnerRow.Cells("OWNERNAME").Value.ToString)
    '        colParams.Add("<Owner Greeting>", ugOwnerRow.Cells("OWNERNAME").Value.ToString.Trim + ":")

    '        colParams.Add("<Owner Address 1>", pOwner.Address.AddressLine1)
    '        If pOwner.Address.AddressLine2 = String.Empty Then
    '            colParams.Add("<Owner Address 2>", pOwner.Address.City + ", " + pOwner.Address.State.TrimEnd + " " + pOwner.Address.Zip)
    '            colParams.Add("<City/State/Zip>", "")
    '        Else
    '            colParams.Add("<Owner Address 2>", pOwner.Address.AddressLine2)
    '            colParams.Add("<City/State/Zip>", pOwner.Address.City + ", " + pOwner.Address.State.TrimEnd + " " + pOwner.Address.Zip)
    '        End If

    '        ' create table of rescinded citations
    '        Dim dt As New DataTable
    '        Dim dr As DataRow
    '        dt.Columns.Add("FACILITY_ID")
    '        dt.Columns.Add("FACILITY")
    '        dt.Columns.Add("STATECITATION")
    '        dt.Columns.Add("CATEGORY")
    '        dt.Columns.Add("CITATIONTEXT")

    '        For Each ugGrandChildRow As Infragistics.Win.UltraWinGrid.UltraGridRow In ugOwnerRow.ChildBands(0).Rows
    '            If ugGrandChildRow.Cells("RESCINDED").Value = True Then
    '                dr = dt.NewRow
    '                dr("FACILITY_ID") = ugGrandChildRow.Cells("FACILITY_ID").Value.ToString
    '                dr("FACILITY") = ugGrandChildRow.Cells("FACILITY").Value.ToString
    '                dr("STATECITATION") = ugGrandChildRow.Cells("STATECITATION").Value.ToString
    '                dr("CATEGORY") = ugGrandChildRow.Cells("CATEGORY").Value.ToString
    '                dr("CITATIONTEXT") = ugGrandChildRow.Cells("CITATIONTEXT").Value.ToString
    '                dt.Rows.Add(dr)
    '            End If
    '        Next

    '        colParams.Add("<User>", MusterContainer.AppUser.Name)

    '        Try
    '            Dim strTempPath As String = TmpltPath & "CAE\" & strTemplateName
    '            Dim ltrGen As New MUSTER.BusinessLogic.pLetterGen
    '            Dim wordApp As Word.Application = MusterContainer.GetWordApp
    '            ltrGen.GenerateOCENFARescindLetter(ugOwnerRow.Cells("OWNER_ID").Value, colParams, strTempPath, doc_path + strDOC_NAME, dt, wordApp)
    '            'ltrGen.CreateInspAnnouncementLetters(ownerID, colParams, strTempPath, Doc_Path + strDOC_NAME, facsForOwner, own, progressBarValueIncrement, MusterContainer.GetWordApp)
    '            UIUtilsGen.SaveDocument(ugOwnerRow.Cells("OWNER_ID").Value, uiutilsgen.EntityTypes.Owner, strDOC_NAME, strDocType, Doc_Path, strDocDesc)
    '            strOCEGeneratedLetterName = strDOC_NAME
    '            wordApp.Visible = True
    '        Catch ex As Exception
    '            Throw ex
    '        End Try
    '        ' need to return true if letter was generated successfully for further processing
    '        Return True
    '    Catch ex As Exception
    '        Throw ex
    '    Finally
    '        SetCursorType(System.Windows.Forms.Cursors.Default)
    '    End Try
    'End Function
    'Friend Function GenerateCAELicenseeLetter(ByVal nFacID As Integer, ByVal nLicenseeID As Integer, ByVal strDocType As String, ByVal strDocName As String, ByVal strDocDesc As String, ByVal strTemplateName As String) As Boolean
    '    Dim pLicensee As New MUSTER.BusinessLogic.pLicensee
    '    Dim pOwner As New MUSTER.BusinessLogic.pOwner
    '    Dim oFacInfo As MUSTER.Info.FacilityInfo
    '    Dim oPersonaInfo As MUSTER.Info.PersonaInfo
    '    Dim oAddressInfo As MUSTER.Info.AddressInfo
    '    Dim oLicenseeAddress As New MUSTER.BusinessLogic.pComAddress
    '    Dim colParams As New Specialized.NameValueCollection
    '    Dim strDOC_NAME As String = String.Empty
    '    Dim strToday As String = String.Empty

    '    Try
    '        If DOC_PATH = "\" Then
    '            Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
    '        End If

    '        strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm"))
    '        strDOC_NAME = "COM_" + strDocName.Trim.ToString + "_" + CStr(Trim(nLicenseeID.ToString)) + "_" + strToday + ".doc"

    '        oFacInfo = pOwner.Facilities.Retrieve(pOwner.OwnerInfo, nFacID, , "FACILITY")
    '        pOwner.Retrieve(oFacInfo.OwnerID)
    '        colParams.Add("<Title>", strDocType.ToString)
    '        colParams.Add("<DATE>", Format(Now, "MMMM dd, yyyy"))

    '        If pOwner.OrganizationID > 0 Then
    '            oPersonaInfo = pOwner.Organization
    '            colParams.Add("<Owner Name>", pOwner.BPersona.Company)
    '        Else
    '            oPersonaInfo = pOwner.Persona
    '            colParams.Add("<Owner Name>", pOwner.BPersona.Title.Trim & IIf(pOwner.BPersona.Title.Length > 0, " ", "") & pOwner.BPersona.FirstName.Trim & IIf(pOwner.BPersona.MiddleName.Length > 0, " ", "") & pOwner.BPersona.MiddleName.Trim & " " & pOwner.BPersona.LastName.Trim & " " & pOwner.BPersona.Suffix)
    '        End If

    '        oAddressInfo = pOwner.Address()
    '        colParams.Add("<Owner Address 1>", oAddressInfo.AddressLine1)
    '        If oAddressInfo.AddressLine2 = String.Empty Then
    '            colParams.Add("<Owner Address 2>", oAddressInfo.City & ", " & oAddressInfo.State.TrimEnd & " " & oAddressInfo.Zip)
    '            colParams.Add("<City/State/Zip>", "")
    '        Else
    '            colParams.Add("<Owner Address 2>", oAddressInfo.AddressLine2)
    '            colParams.Add("<City/State/Zip>", oAddressInfo.City & ", " & oAddressInfo.State.TrimEnd & " " & oAddressInfo.Zip)
    '        End If

    '        oAddressInfo = pOwner.Facilities.FacilityAddress
    '        colParams.Add("<Facility Name>", oFacInfo.Name)
    '        colParams.Add("<Facility Address>", oAddressInfo.AddressLine1 & " " & IIf(oAddressInfo.AddressLine2 = String.Empty, "", oAddressInfo.AddressLine2))
    '        colParams.Add("<FacilityCityStateZip>", oAddressInfo.City.TrimEnd & ", " & oAddressInfo.State.TrimEnd & " " & oAddressInfo.Zip.TrimEnd)
    '        colParams.Add("<I.D. #>", oFacInfo.ID.ToString)

    '        pLicensee.Retrieve(nLicenseeID)
    '        colParams.Add("<Licensee Name>", pLicensee.Licensee_name)
    '        colParams.Add("<Licensee Greeting>", "Dear " & pLicensee.Licensee_name & " :")
    '        oLicenseeAddress.GetAddressByType(0, 0, nLicenseeID, 0)

    '        colParams.Add("<Licensee Address1>", oLicenseeAddress.AddressLine1)
    '        If oLicenseeAddress.AddressLine2 = String.Empty Then
    '            colParams.Add("<Licensee Address2>", oLicenseeAddress.City & ", " & oLicenseeAddress.State.TrimEnd & " " & oLicenseeAddress.Zip)
    '            colParams.Add("<City/State/Zip>", "")
    '        Else
    '            colParams.Add("<Licensee Address2>", oLicenseeAddress.AddressLine2)
    '            colParams.Add("<City/State/Zip>", oLicenseeAddress.City & ", " & oLicenseeAddress.State.TrimEnd & " " & oLicenseeAddress.Zip)
    '        End If
    '        colParams.Add("<User>", MusterContainer.AppUser.Name)

    '        Try
    '            Dim strTempPath As String = TmpltPath & "CAE\" & strTemplateName
    '            UIUtilsGen.CreateAndSaveDocument("CAE", nLicenseeID, 31, DOC_PATH, strDOC_NAME, strDocType, strTempPath, Doc_Path, strDocDesc, colParams)
    '        Catch ex As Exception
    '            Throw ex
    '        End Try
    '    Catch ex As Exception
    '        Throw ex
    '    Finally
    '        SetCursorType(System.Windows.Forms.Cursors.Default)
    '    End Try
    'End Function
#End Region
#Region "Fees Letters"

    Friend Function GenerateFeesLetter(ByVal strDocType As String, ByVal strDocName As String, ByVal strDocDesc As String, ByVal strTemplateName As String, ByRef pOwner As MUSTER.BusinessLogic.pOwner, Optional ByVal LateCertID As Int64 = 0, Optional ByVal strFacilities As String = "", Optional ByVal certNumber As String = "") As Boolean
        Dim strContactName As String
        Dim strOwnerName As String
        Dim bolNeedCapDocs As Boolean = False
        Dim strDOC_NAME As String = String.Empty
        Dim strToday As String = String.Empty

        Dim oOwnerInfo As MUSTER.Info.OwnerInfo
        Dim oPersonaInfo As MUSTER.Info.PersonaInfo
        Dim oAddressInfo As MUSTER.Info.AddressInfo

        Dim oContact As New MUSTER.BusinessLogic.pContactDatum
        Dim oContactInfo As New MUSTER.Info.ContactDatumInfo
        Dim oFeesLateCert As New MUSTER.BusinessLogic.pFeeLateFee

        Dim colParams As New Specialized.NameValueCollection
        nModuleID = UIUtilsGen.ModuleID.Fees
        Try
            If DOC_PATH = "\" Then
                Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
            End If
            strContactName = ""

            strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm"))
            strDOC_NAME = "FEES_" + strDocName.Trim.ToString + "_" + CStr(Trim(pOwner.ID)) + "_" + strToday + ".doc"

            'Build NameValueCollection with Tags and Values.
            oOwnerInfo = pOwner.Retrieve(pOwner.ID, "SELF")

            colParams.Add("<Letter Date>", Today.ToString("MMMM d, yyyy"))
            colParams.Add("<DueDate>", Today.AddDays(30).ToString("MMMM d, yyyy"))


            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Contact name
            Dim dtContacts As DataTable
            Dim drRow As DataRow
            Dim strXHContactName As String = String.Empty
            Dim strXLContactName As String = String.Empty
            Dim strRXLContactName As String = String.Empty

            dtContacts = GetXHAndXLContacts(pOwner.ID, UIUtilsGen.EntityTypes.LUST_Event, UIUtilsGen.ModuleID.Fees, pOwner.ID, 1)
            If dtContacts.Rows.Count > 0 Then
                For Each drRow In dtContacts.Rows
                    If drRow("EntityID") = pOwner.ID Then
                        If drRow("Type") = EnumContactType.XH Then
                            strXHContactName = drRow("CONTACT_Name")
                            colParams.Add("<Company Address>", drRow("Address_One").ToString & " " & IIf(drRow("Address_Two").ToString = String.Empty, "", "; " & drRow("Address_Two").ToString))
                            colParams.Add("<City, State, Zip>", drRow("City").ToString & ", " & drRow("State").ToString & " " & drRow("Zip").ToString)
                        ElseIf drRow("Type") = EnumContactType.XL Then
                            strXLContactName = drRow("CONTACT_Name")
                        End If
                    ElseIf drRow("EntityID") = pOwner.ID Then
                        If drRow("Type") = EnumContactType.XL Then
                            strRXLContactName = drRow("CONTACT_Name")
                        End If
                    End If
                Next

                If strXHContactName <> String.Empty And strXLContactName <> String.Empty Then
                    colParams.Add("<Company Name>", strXHContactName)
                    colParams.Add("<Owner Contact>", strXLContactName)
                    colParams.Add("<Salutation>", strXLContactName)
                ElseIf (strXHContactName = String.Empty And strXLContactName <> String.Empty) Then
                    colParams.Add("<Owner Contact>", strXLContactName)
                    colParams.Add("<Salutation>", strXLContactName)
                ElseIf strXHContactName <> String.Empty And strXLContactName = String.Empty Then
                    colParams.Add("<Company Name>", strXHContactName)
                    colParams.Add("<Owner Contact>", "")
                    colParams.Add("<Salutation>", strXHContactName)
                ElseIf strRXLContactName <> String.Empty Then
                    colParams.Add("<Owner Contact>", strRXLContactName)
                    colParams.Add("<Salutation>", strRXLContactName)
                End If
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If (strXHContactName = String.Empty And strXLContactName = String.Empty) Or (strXHContactName = String.Empty) Then
                If oOwnerInfo.OrganizationID > 0 Then
                    oPersonaInfo = pOwner.Organization
                    colParams.Add("<Company Name>", pOwner.BPersona.Company)
                    If strXLContactName = String.Empty And strRXLContactName = String.Empty Then
                        colParams.Add("<Salutation>", pOwner.BPersona.Company.Trim)
                    End If
                    colParams.Add("<OWNER NAME>", pOwner.BPersona.Company)
                Else
                    oPersonaInfo = pOwner.Persona
                    colParams.Add("<Company Name>", pOwner.BPersona.FirstName.Trim & IIf(pOwner.BPersona.MiddleName.Trim.Length > 0, " ", "") & pOwner.BPersona.MiddleName.Trim & IIf(pOwner.BPersona.LastName.Trim.Length > 0, " ", "") & pOwner.BPersona.LastName.Trim & IIf(pOwner.BPersona.Suffix.Trim.Length > 0, " ", "") & pOwner.BPersona.Suffix.Trim)
                    If strXLContactName = String.Empty And strRXLContactName = String.Empty Then
                        colParams.Add("<Salutation>", pOwner.BPersona.FirstName.Trim & IIf(pOwner.BPersona.MiddleName.Trim.Length > 0, " ", "") & pOwner.BPersona.MiddleName.Trim & IIf(pOwner.BPersona.LastName.Trim.Length > 0, " ", "") & pOwner.BPersona.LastName.Trim & IIf(pOwner.BPersona.Suffix.Trim.Length > 0, " ", "") & pOwner.BPersona.Suffix.Trim)
                    End If
                End If
            End If
            If dtContacts.Rows.Count = 0 Then
                colParams.Add("<Owner Contact>", "")
            End If

            If strXHContactName = String.Empty Then
                oAddressInfo = pOwner.Address()
                colParams.Add("<Company Address>", oAddressInfo.AddressLine1 & " " & IIf(oAddressInfo.AddressLine2 = String.Empty, "", "; " & oAddressInfo.AddressLine2))
                colParams.Add("<City, State, Zip>", oAddressInfo.City & ", " & oAddressInfo.State.TrimEnd & " " & oAddressInfo.Zip)
            End If

            colParams.Add("<Due Date>", DateAdd(DateInterval.Day, 30, Now()).ToShortDateString)
            If strFacilities.Length > 0 Then
                colParams.Add("<FACILITIES>", strFacilities)
            Else
                colParams.Add("RE: ", String.Empty)
            End If


            oFeesLateCert.Retrieve(LateCertID)


            If oFeesLateCert.CertLetterNumber > "" Then
                colParams.Add("<Cert Mail Number>", oFeesLateCert.CertLetterNumber)
            Else
                If colParams.Item("<Cert Mail Number>") = String.Empty Then
                    colParams.Add("<Cert Mail Number>", certNumber.Substring(0, 24))
                End If

            End If


            Try
                Dim strTempPath As String = TmpltPath & "Fees\" & strTemplateName
                UIUtilsGen.CreateAndSaveDocument("Fees", pOwner.ID, UIUtilsGen.EntityTypes.Owner, DOC_PATH, strDOC_NAME, strDocType, strTempPath, Doc_Path, strDocDesc, nModuleID, colParams, 0, 0, 0)

                Return True
            Catch ex As Exception
                Throw ex
            End Try

        Catch ex As Exception
            Throw ex
        Finally


            SetCursorType(System.Windows.Forms.Cursors.Default)
        End Try
    End Function

#End Region
#Region "Generic Letters"
    Friend Function GenerateGenericLetter(ByVal ModuleID As Integer, ByVal Title As String, ByVal BodyData As String, ByVal Cols As Int16, Optional ByRef IsDraft As Boolean = True, Optional ByVal ModuleName As String = "Global", Optional ByRef EntityID As Int64 = 0, Optional ByVal EntityType As Int64 = 0, Optional ByVal DocNamePrefix As String = "FAC_GEN_", Optional ByVal DocDescription As String = "", Optional ByVal DocType As String = "", Optional ByVal TableFormat As Int16 = 35, Optional ByVal bApplyBorders As Boolean = False, Optional ByVal bApplyShading As Boolean = False, Optional ByVal bApplyFont As Boolean = True, Optional ByVal bApplyColor As Boolean = False, Optional ByVal bApplyHeader As Boolean = True, Optional ByVal bApplyLastRow As Boolean = False, Optional ByVal bApplyFirstCol As Boolean = False, Optional ByVal bApplyLastCol As Boolean = False, Optional ByVal bApplyAutofit As Boolean = True, Optional ByVal attachDocumentInfo As String = "", Optional ByVal eventID As Int64 = 0, Optional ByVal eventSequence As Integer = 0, Optional ByVal eventType As Integer = 0) As Boolean

        Dim bolNeedCapDocs As Boolean = False
        Dim strDOC_NAME As String = String.Empty
        Dim strToday As String = String.Empty
        Dim i As Int16
        Dim strYes As String
        Dim strNo As String
        Dim strOther As String
        Dim colParams As New Specialized.NameValueCollection
        Dim strDocPath As String
        Dim tmpDate As Date
        Dim oWordApp As Word.Application
        Dim fileName As Object
        Dim aDoc As Word.Document
        Dim areadOnly As Object = True
        Dim isVisible As Object = True
        Dim confirmConversions As Object = False
        Dim addToRecentFiles As Object = False
        Dim revert As Object = False
        Dim missing As Object = System.Reflection.Missing.Value

        Try

            If DOC_PATH = "\" Then
                Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
            End If

            strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm")) + "_" + CStr(Format(Now, "ss"))
            strDOC_NAME = DocNamePrefix + CStr(Trim(EntityID.ToString)) + "_" + strToday + ".doc"
            colParams.Add("<DATE>", Format(Now, "MMMM d, yyyy"))
            colParams.Add("<TITLE>", Title)
            colParams.Add("<DATA>", BodyData)

            Dim ltrGen As New MUSTER.BusinessLogic.pLetterGen
            Dim strTempPath As String = System.IO.Path.GetTempPath

            If IsDraft Then
                strDocPath = strTempPath
            Else
                strDocPath = DOC_PATH
            End If

            oWordApp = MusterContainer.GetWordApp

            If Not oWordApp Is Nothing Then
                ltrGen.CreateGenericLetter(ModuleName, strDOC_NAME, colParams, Cols, TmpltPath & "Global\Generic.doc", strDocPath & strDOC_NAME, oWordApp, TableFormat, bApplyBorders, bApplyShading, bApplyFont, bApplyColor, bApplyHeader, bApplyLastRow, bApplyFirstCol, bApplyLastCol, bApplyAutofit, attachDocumentInfo)
                'Delay()
                UIUtilsGen.Delay(, 1)


                ' Make word visible
                oWordApp.Visible = True

                If IsDraft Then
                    ' Open the document that was chosen by the dialog
                    aDoc = oWordApp.Documents.Open(strDocPath & strDOC_NAME, confirmConversions, areadOnly, addToRecentFiles, missing, missing, revert, missing, missing, missing, missing, isVisible)

                Else
                    UIUtilsGen.SaveDocument(EntityID, EntityType, strDOC_NAME, DocType, DOC_PATH, DocDescription, ModuleID, eventID, eventSequence, eventType)
                End If
            End If



        Catch ex As Exception
            Throw ex
        Finally
            SetCursorType(System.Windows.Forms.Cursors.Default)
        End Try
    End Function
#End Region

#Region "External Events"
    Private Sub CertifiedMail(ByVal strCertMail As String) Handles certMail.evtCertifiedMail
        strCertifiedMail = strCertMail
    End Sub
#End Region
#Region "Envelopes And Letters"
    Public Function CreateEnvelopes(ByVal strAddress As String, ByVal strModule As String, ByVal nEntityID As Integer)
        Dim colParams As New Specialized.NameValueCollection
        Dim strTempPath As String
        Dim ltrGen As New MUSTER.BusinessLogic.pLetterGen
        Dim strDOC_NAME As String = String.Empty
        Dim strToday As String = String.Empty
        Dim sAddress() As String
        Try
            If DOC_PATH = "\" Then
                Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
            End If

            strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm"))
            strDOC_NAME = strModule + "Envelope" + "_" + CStr(Trim(nEntityID.ToString)) + "_" + strToday + ".doc"
            sAddress = strAddress.Split(",")
            colParams.Add("<NAME>", sAddress(0))
            colParams.Add("<ADDRESS1>", sAddress(1))
            If sAddress(2) = String.Empty Then
                colParams.Add("<ADDRESS2>", sAddress(3) + " " + sAddress(4) + " - " + sAddress(5))
                colParams.Add("<CITYSTATEZIP>", String.Empty)
            Else
                colParams.Add("<ADDRESS2>", sAddress(2))
                colParams.Add("<CITYSTATEZIP>", sAddress(3) + " " + sAddress(4) + " - " + sAddress(5))
            End If
            strTempPath = TmpltPath & "Global\EnvelopeTemplate.doc"
            Dim owordApp As Word.Application = MusterContainer.GetWordApp
            If Not owordApp Is Nothing Then
                ltrGen.CreateEnvelope("Envelope", colParams, strTempPath, String.Empty, owordApp)
                'ltrGen.CreateLetter("Global", "Envelope", colParams, strTempPath, Doc_Path & strDOC_NAME, owordApp)
                owordApp.Visible = True
            End If
            owordApp = Nothing

        Catch ex As Exception
            Throw ex

        End Try
    End Function
    Public Function CreateLabels(ByVal strAddress As String, ByVal strModule As String, ByVal nEntityID As Integer)
        Dim colParams As New Specialized.NameValueCollection
        Dim strTempPath As String
        Dim ltrGen As New MUSTER.BusinessLogic.pLetterGen
        Dim strDOC_NAME As String = String.Empty
        Dim strToday As String = String.Empty
        Dim sAddress() As String
        Try
            If DOC_PATH = "\" Then
                Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
            End If

            strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm"))
            strDOC_NAME = strModule + "Label" + "_" + CStr(Trim(nEntityID.ToString)) + "_" + strToday + ".doc"
            sAddress = strAddress.Split(",")
            colParams.Add("<NAME>", sAddress(0))
            colParams.Add("<ADDRESS1>", sAddress(1))
            If sAddress(2) = String.Empty Then
                colParams.Add("<ADDRESS2>", sAddress(3) + " " + sAddress(4) + " - " + sAddress(5))
                colParams.Add("<CITYSTATEZIP>", String.Empty)
            Else
                colParams.Add("<ADDRESS2>", sAddress(2))
                colParams.Add("<CITYSTATEZIP>", sAddress(3) + " " + sAddress(4) + " - " + sAddress(5))
            End If
            strTempPath = TmpltPath & "Global\Labels.doc"
            Dim owordApp As Word.Application = MusterContainer.GetWordApp
            If Not owordApp Is Nothing Then
                ltrGen.CreateLabels("Global", "Labels", colParams, strTempPath, String.Empty, owordApp)
                'ltrGen.CreateLetter("Global", "Envelope", colParams, strTempPath, Doc_Path & strDOC_NAME, owordApp)
                owordApp.Visible = True
            End If
            owordApp = Nothing
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region
End Class
