'MR - 10/10/2004
Imports System.IO
Imports Word
Imports Word.ApplicationClass


'Imports InfoRepository
Friend Class CAP_Letters
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
    '                                 
    '-------------------------------------------------------------------------------
    Inherits LetterGenerator
    Private WithEvents WordApp As Word.Application
    Private oPara As Word.Paragraph
    Dim DestDoc As Word.Document

    Private WithEvents ug As New Infragistics.Win.UltraWinGrid.UltraGrid
    Private ugRow, ugChildRow, ugGrandChildRow, ugGreatGrandChildRow As Infragistics.Win.UltraWinGrid.UltraGridRow

    Private slCAPAnnualSummaryDocs As New SortedList

    Private strCAPAnnualMailMergeNoCalDocName, strCAPAnnualMailMergeCalDocName As String



    Public Enum CapAnnualMode
        StaticByYear = 0
        CurrentSummary = 1
    End Enum



    Sub SetupSystemToGenerateCAPMonthly(ByVal prompt As Boolean, ByVal strOwnerName As String, ByVal fac_id As Integer)

        Try

            Dim pass As Boolean = False
            'check rights for Tank & Pipe
            If Not _container.AppUser.HasAccess(CType(UIUtilsGen.ModuleID.CAPProcess, Integer), MusterContainer.AppUser.UserKey, UIUtilsGen.EntityTypes.Tank) Then
                MessageBox.Show("You do not have Rights to Monthly CAP Processing")
                Exit Sub
            ElseIf Not _container.AppUser.HasAccess(CType(UIUtilsGen.ModuleID.CAPProcess, Integer), MusterContainer.AppUser.UserKey, UIUtilsGen.EntityTypes.Pipe) Then
                MessageBox.Show("You do not have Rights to Monthly CAP Processing")
                Exit Sub
            End If

            Dim strYear As DateTime
            Dim ownerName As String = strOwnerName


            strYear = InputBox("Enter DATE (Month/day/Year) for Monthly CAP report", , Today)

            If Not IsDate(strYear) Then
                MsgBox("Invalid entry")
                Exit Sub
            End If

            If prompt Then
                ownerName = InputBox("Enter Owner Name (Leave Blank for Full Monthly CAP Report)", , String.Empty)
            End If

            _container.Cursor = Cursors.WaitCursor

            strYear = New Date(strYear.Year, strYear.Month, 1)

            pass = GenerateCAPMonthly(strYear, _container.pOwn, ownerName)

            If pass Then

                Dim frmReport As ReportDisplay

                frmReport = New ReportDisplay
                frmReport.MdiParent = _container

                Dim ownID As Integer = 0

                If Not _container.pOwn Is Nothing Then
                    ownID = _container.pOwn.ID
                End If

                Try

                    If prompt Then
                        strOwnerName = ownerName
                    End If

                    If strOwnerName.Length = 0 Then
                        strOwnerName = " "
                    End If

                    ' frmReport.Show()
                    frmReport.GenerateReport("CAP Monthly Summary Report", New Object() {strYear.Year, IIf(ownID = 0, DBNull.Value, ownID), strOwnerName, strYear.Month}, "PROCESSED")

                Catch ex As Exception
                    Dim MyErr As ErrorReport
                    MyErr = New ErrorReport(New Exception("Error in loading reports form : " & vbCrLf & ex.Message, ex))
                    MyErr.ShowDialog()
                End Try
            End If


        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            _container.Cursor = Cursors.Default
        End Try

    End Sub


    Friend Function GenerateCAPMonthly(ByVal processingMonthYear As Date, ByVal pOwn As MUSTER.BusinessLogic.pOwner, Optional ByVal ownerName As String = "") As Boolean
        Dim ds As DataSet

        Dim processingYear = processingMonthYear.Year
        Dim processingMonth = processingMonthYear.Month

        Dim oldText As String = _container.ActiveForm.Text



        Dim dsRelOwnerFac, dsRelFacTank, dsRelTankPipe As DataRelation
        Dim dtOwner, dtFac, dtTank, dtPipe As DataTable
        Dim drOwner, drFac, drTank, drPipe As DataRow ' using different variables for less confusion
        Dim drdsOwner, drdsFac, drdsTank, drdsPipe As DataRow ' for looping through the tables
        Dim drdsFacs() As DataRow
        Dim dtProcessingStart, dtProcessingEnd As Date
        'Dim strTestReqDocName As String = ""
        'Dim strAssistDocName As String = ""
        ' Dim strTemplate As String = ""
        ' Dim strAssistTemplate As String = ""

        Dim bolDeleteFilesCreated As Boolean = False

        Dim headingText As String = ""
        Dim dt, dt1 As Date

        Dim bolEnteredTextToTestReqLetter As Boolean = False
        Dim bolEnteredTextToAssistLetter As Boolean = False

        Dim bolAddedSectionForOwner As Boolean = False
        Dim bolAddedFacilityDetail As Boolean = False

        'Dim docAssist, docTestReq As Word.Document

        Dim alTnkLastTCPDate As New ArrayList
        Dim alTnkLinedDate As New ArrayList
        Dim alTnkTTDate As New ArrayList
        Dim alTnkICExpiresDate As New ArrayList
        Dim alTnkSpillDate As New ArrayList
        Dim alTnkOverfillDate As New ArrayList
        Dim alTnkSecondaryDate As New ArrayList
        Dim alTnkElectronicDate As New ArrayList
        Dim alTnkATGDate As New ArrayList

        Dim alPipeCPDate As New ArrayList
        Dim alPipeALLDDate As New ArrayList
        Dim alPipeTermCPDate As New ArrayList
        Dim alPipeLineUSDate As New ArrayList
        Dim alPipeLinePressDate As New ArrayList
        Dim alPipeSheerDate As New ArrayList
        Dim alPipeSecondaryDate As New ArrayList
        Dim alPipeElectronicDate As New ArrayList
        Dim strOwnerIDs As String = String.Empty
        Dim drCalInfo As DataRow

        Try

            ' datatables to maintain the records of the information to populate calendar
            Dim dtCalInfo = New DataTable

            dtCalInfo.Columns.Add("FACILITY_ID", GetType(Integer))
            dtCalInfo.Columns.Add("MONTH", GetType(Integer))
            dtCalInfo.Columns.Add("INFO", GetType(String))
            dtCalInfo.Columns.Add("NAME", GetType(String))
            dtCalInfo.Columns.Add("CITY", GetType(String))


            If DOC_PATH = "\" Then
                MsgBox("Document Path Unspecified. Please give the path before generating the letter.")
                Exit Function
            End If

            If Date.Compare(processingMonthYear, CDate("01/01/0001")) = 0 Then
                processingMonthYear = CDate(Today.Month.ToString + "/1/" + Today.Year.ToString)
            ElseIf processingMonthYear.Day <> 1 Then
                processingMonthYear = CDate(processingMonthYear.Month.ToString + "/1/" + processingMonthYear.Year.ToString)
            End If

            ' dtProcessingstart = 2 months from processingMonthYear
            dtProcessingStart = DateAdd(DateInterval.Month, 2, processingMonthYear)
            ' add a month and substract a day to get the last day of the month
            dtProcessingEnd = DateAdd(DateInterval.Month, 1, dtProcessingStart)
            dtProcessingEnd = DateAdd(DateInterval.Day, -1, dtProcessingEnd)

            '''To avoid duplicate creation of Letters.
            'strTestReqDocName = "REG_CAP_PROCESSING_RPT_" + processingMonthYear.Month.ToString + "-" + processingMonthYear.Year.ToString + ownerName + ".doc"
            'strAssistDocName = "REG_Compliance_Assistance_Letter_" + processingMonthYear.Month.ToString + "-" + processingMonthYear.Year.ToString + ownerName + ".doc"

            'If FileExists(DOC_PATH + strTestReqDocName) OrElse FileExists(DOC_PATH + strAssistDocName) Then
            'if MsgBox("CAP Processing Report for " + IIf(ownerName.Length > 0, String.Format(" Owner '{0}' for", ownerName), String.Empty) + processingMonthYear.Month.ToString + "-" + processingMonthYear.Year.ToString + "  has been created already. Would like to regenerate this report? ", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then

            'File.Delete(DOC_PATH + strTestReqDocName)
            'File.Delete(DOC_PATH + strAssistDocName)

            ' Threading.Thread.Sleep(2000)

            ' Else
            '    Exit Sub
            ' End If

            'End If


            'dtEnd = DateTime.Now
            'ts = dtEnd.Subtract(dtStart)
            'strTime += "check path and processing month/year: " + ts.ToString

            'dtStart = DateTime.Now

            ds = pOwn.RunSQLQuery("EXEC spSelCapProcess 0, '" + dtProcessingStart.ToShortDateString + "', '" + dtProcessingEnd.ToShortDateString + "'")


            If ds.Tables(0).Rows.Count > 0 Then ' if owner has no rows, facility / tanks / pipes will not have rows


                ''' create file for Test Req
                'strTemplate = TmpltPath + "CAP\CapMonthlyTestingReqHeading.doc"
                ' If Not System.IO.File.Exists(strTemplate) Then
                'MsgBox("Template(" + strTemplate + " not found")
                ' Exit Sub
                'End If

                'strTemplate = TmpltPath + "CAP\CapMonthlyTestReq.doc"
                'If Not System.IO.File.Exists(strTemplate) Then
                '   MsgBox("Template(" + strTemplate + " not found")
                '  Exit Sub
                ' End If
                'System.IO.File.Copy(strTemplate, doc_path + strTestReqDocName)

                ' create file for Assist
                'strAssistTemplate = TmpltPath + "CAP\CapMonthlyAssistance.doc"
                'If Not System.IO.File.Exists(strAssistTemplate) Then
                '   MsgBox("Template(" + strAssistTemplate + " not found")
                '  Exit Sub
                'End If
                'System.IO.File.Copy(strAssistTemplate, doc_path + strAssistDocName)


                'WordApp = MusterContainer.GetWordApp


                ' docAssist = WordApp.Documents.Open(doc_path + strAssistDocName)
                ' docTestReq = WordApp.Documents.Open(doc_path + strTestReqDocName)

                ''' enter date in footer
                'docTestReq.Activate()
                'With docTestReq
                '   .ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter
                '  .Application.Selection.Find.Execute(FindText:="<Date>", ReplaceWith:=Now.Date.ToShortDateString, Replace:=Word.WdReplace.wdReplaceAll)
                ' .ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument
                'End With


                'headingText = GetCapMonthlyNoticeOfTestReqHeading(WordApp)

                ' datatables to maintain the records of the tanks and pipes whose dates are to be rolledover
                dtOwner = New DataTable
                dtFac = New DataTable
                dtTank = New DataTable
                dtPipe = New DataTable

                dtOwner.Columns.Add("OWNER_ID", GetType(Integer))

                dtFac.Columns.Add("OWNER_ID", GetType(Integer))
                dtFac.Columns.Add("FACILITY_ID", GetType(Integer))

                dtTank.Columns.Add("FACILITY_ID", GetType(Integer))
                dtTank.Columns.Add("TANK ID", GetType(Integer))
                dtTank.Columns.Add("CP DATE", GetType(Date))
                dtTank.Columns.Add("LI INSPECTED", GetType(Date))
                dtTank.Columns.Add("TT DATE", GetType(Date))

                'added by Hua Cao 11/12/2008
                dtTank.Columns.Add("DateSpillPreventionInstalled", GetType(Date))
                dtTank.Columns.Add("DateSpillPreventionLastTested", GetType(Date))
                dtTank.Columns.Add("DateOverfillPreventionLastInspected", GetType(Date))
                dtTank.Columns.Add("DateSecondaryContainmentLastInspected", GetType(Date))
                dtTank.Columns.Add("DateElectronicDeviceInspected", GetType(Date))
                dtTank.Columns.Add("DateATGLastInspected", GetType(Date))
                dtTank.Columns.Add("DateOverfillPreventionInstalled", GetType(Date))


                dtPipe.Columns.Add("TANK ID", GetType(Integer))
                dtPipe.Columns.Add("PIPE ID", GetType(Integer))
                dtPipe.Columns.Add("CP DATE", GetType(Date))
                dtPipe.Columns.Add("TERM CP TEST", GetType(Date))
                dtPipe.Columns.Add("ALLD_TEST_DATE", GetType(Date))
                dtPipe.Columns.Add("TT DATE", GetType(Date))
                'added by Hua Cao 11/12/2008
                dtPipe.Columns.Add("DateSheerValueTest", GetType(Date))
                dtPipe.Columns.Add("DateSecondaryContainmentInspect", GetType(Date))
                dtPipe.Columns.Add("DateElectronicDeviceInspect", GetType(Date))

                'With WordApp

                '            .Visible = True

                'dtStart = DateTime.Now
                'Owner loop for CAP monthly report

                Dim cnt As Integer = ds.Tables(0).Rows.Count

                For i As Integer = 0 To cnt - 1 ' owner

                    If String.Format("{0}   Preparing Monthly CAP: {1}% ", oldText, Int((((i + 1) / cnt) * 100))) <> _container.Text Then

                        _container.Text = String.Format("{0}   Preparing Monthly CAP: {1}% ", oldText, Int((((i + 1) / cnt) * 100)))

                    End If

                    bolEnteredTextToTestReqLetter = False


                    'For Each drdsOwner In ds.Tables(0).Rows ' owner
                    drdsOwner = ds.Tables(0).Rows(i)

                    dtCalInfo.Rows.Clear()

                    If ownerName = String.Empty OrElse drdsOwner("OWNERNAME").ToString.ToUpper.IndexOf(ownerName.ToUpper) > -1 Then


                        drOwner = dtOwner.NewRow
                        drOwner("OWNER_ID") = drdsOwner("OWNER_ID")


                        'With docAssist

                        bolAddedSectionForOwner = False


                        drdsFacs = ds.Tables(1).Select("OWNER_ID = " + drdsOwner("OWNER_ID").ToString)


                        'Facility Loop for CAP monthly report
                        For j As Integer = 0 To drdsFacs.Length - 1 ' facility
                            'For Each drdsFac In ds.Tables(1).Select("OWNER_ID = " + drdsOwner("OWNER_ID").ToString) ' facility

                            drdsFac = drdsFacs(j)

                            drFac = dtFac.NewRow
                            drFac("OWNER_ID") = drdsOwner("OWNER_ID")
                            drFac("FACILITY_ID") = drdsFac("FACILITY_ID")

                            bolAddedFacilityDetail = False

                            alTnkLastTCPDate = New ArrayList
                            alTnkLinedDate = New ArrayList
                            alTnkTTDate = New ArrayList
                            alTnkICExpiresDate = New ArrayList
                            alTnkSpillDate = New ArrayList
                            alTnkOverfillDate = New ArrayList
                            alTnkSecondaryDate = New ArrayList
                            alTnkElectronicDate = New ArrayList
                            alTnkATGDate = New ArrayList

                            alPipeCPDate = New ArrayList
                            alPipeALLDDate = New ArrayList
                            alPipeTermCPDate = New ArrayList
                            alPipeLineUSDate = New ArrayList
                            alPipeLinePressDate = New ArrayList
                            alPipeSheerDate = New ArrayList
                            alPipeSecondaryDate = New ArrayList
                            alPipeElectronicDate = New ArrayList

                            '   With docTestReq
                            '  docTestReq.Activate()
                            'Tank loop for the CAP monthly report

                            For Each drdsTank In ds.Tables(2).Select("FACILITY_ID = " + drdsFac("FACILITY_ID").ToString) ' tank

                                If drdsFac("FACILITY_ID").ToString = "3161" Then
                                    Dim test As String
                                    test = "0"
                                End If

                                drTank = dtTank.NewRow
                                drTank("FACILITY_ID") = drdsFac("FACILITY_ID")
                                drTank("TANK ID") = drdsTank("TANK ID")


                                If Not drdsTank("STATUS") Is DBNull.Value Then
                                    If drdsTank("STATUS").ToString.IndexOf("Currently In Use") > -1 Or drdsTank("STATUS").ToString.IndexOf("Temporarily Out of Service Indefinitely") > -1 Then


                                        If drdsTank("Substance") <> "Used Oil" And drdsTank("SMALLDELIVERY") = 0 And drdsTank("STATUS").ToString.IndexOf("Temporarily Out of Service Indefinitely") <= -1 Then


                                            'DateSpillPreventionLastTested
                                            If Not drdsTank("DateSpillPreventionLastTested") Is DBNull.Value Then

                                                dt = drdsTank("DateSpillPreventionLastTested")
                                                dt = dt.Date
                                                dt = DateAdd(DateInterval.Year, 1, dt)

                                                If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                    If Not alTnkSpillDate.Contains(dt.ToShortDateString) Then

                                                        alTnkSpillDate.Add(dt.ToShortDateString)

                                                        drCalInfo = dtCalInfo.NewRow
                                                        drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                        drCalInfo("NAME") = drdsFac("NAME")
                                                        drCalInfo("CITY") = drdsFac("CITY")
                                                        drCalInfo("MONTH") = dt.Month
                                                        drCalInfo("INFO") = "Testing of spill containment buckets must be accomplished by " + dt.ToShortDateString + "." + vbCrLf + _
                                                                            "Please update my records to reflect that this test was accomplished on _______________."
                                                        dtCalInfo.Rows.Add(drCalInfo)

                                                    End If

                                                    bolEnteredTextToTestReqLetter = True
                                                    drdsTank("MODIFIED") = True
                                                    drdsFac("MODIFIED") = True
                                                    drdsOwner("MODIFIED") = True

                                                End If
                                            Else

                                                alTnkSpillDate.Add(CDate("1/1/1900"))

                                                drCalInfo = dtCalInfo.NewRow
                                                drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                drCalInfo("NAME") = drdsFac("NAME")
                                                drCalInfo("CITY") = drdsFac("CITY")
                                                drCalInfo("MONTH") = dt.Month
                                                drCalInfo("INFO") = "Last testing date of spill containment buckets is unknown." + vbCrLf + _
                                                                    "Please update my records to reflect that this test was accomplished on _______________."
                                                dtCalInfo.Rows.Add(drCalInfo)


                                            End If

                                            'DateOverfillPreventionLastInspected
                                            If Not drdsTank("DateOverfillPreventionLastInspected") Is DBNull.Value Then

                                                dt = drdsTank("DateOverfillPreventionLastInspected")
                                                dt = dt.Date
                                                dt = DateAdd(DateInterval.Year, 1, dt)

                                                If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then

                                                    If Not alTnkOverfillDate.Contains(dt.ToShortDateString) Then
                                                        alTnkOverfillDate.Add(dt.ToShortDateString)

                                                        drCalInfo = dtCalInfo.NewRow
                                                        drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                        drCalInfo("NAME") = drdsFac("NAME")
                                                        drCalInfo("CITY") = drdsFac("CITY")
                                                        drCalInfo("MONTH") = dt.Month
                                                        drCalInfo("INFO") = "Inspection of overfill prevention must be accomplished by " + dt.ToShortDateString + "." + vbCrLf + _
                                                                            "Please update my records to reflect that this test was accomplished on _______________."
                                                        dtCalInfo.Rows.Add(drCalInfo)

                                                    End If

                                                    bolEnteredTextToTestReqLetter = True
                                                    drdsTank("MODIFIED") = True
                                                    drdsFac("MODIFIED") = True
                                                    drdsOwner("MODIFIED") = True
                                                End If
                                            Else
                                                alTnkOverfillDate.Add(CDate("1/1/1900"))

                                                drCalInfo = dtCalInfo.NewRow
                                                drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                drCalInfo("NAME") = drdsFac("NAME")
                                                drCalInfo("CITY") = drdsFac("CITY")
                                                drCalInfo("MONTH") = dt.Month
                                                drCalInfo("INFO") = "Last inspection date of overfill prevention is unknown." + vbCrLf + _
                                                                    "Please update my records to reflect that this test was accomplished on _______________."
                                                dtCalInfo.Rows.Add(drCalInfo)
                                            End If

                                            'If drdsTank("Tank_LD_Num") = 343 And drdsTank("STATUS").ToString.IndexOf("Currently In Use") > -1 Then


                                            ''DateSecondaryContainmentLastInspected
                                            ' If Not drdsTank("DateSecondaryContainmentLastInspected") Is DBNull.Value Then

                                            'dt = drdsTank("DateSecondaryContainmentLastInspected")
                                            'dt = dt.Date
                                            'dt = DateAdd(DateInterval.Year, 1, dt)

                                            'If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then

                                            'If Not alTnkSecondaryDate.Contains(dt.ToShortDateString) Then

                                            'alTnkSecondaryDate.Add(dt.ToShortDateString)

                                            'drCalInfo = dtCalInfo.NewRow
                                            'drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                            'drCalInfo("NAME") = drdsFac("NAME")
                                            'drCalInfo("CITY") = drdsFac("CITY")
                                            'drCalInfo("MONTH") = dt.Month
                                            'drCalInfo("INFO") = "Inspection of the tank secondary containment must be accomplished by " + dt.ToShortDateString + "." + vbCrLf + _
                                            '                        "Please update my records to reflect that this test was accomplished on _______________."
                                            'dtCalInfo.Rows.Add(drCalInfo)

                                            'End If

                                            'bolEnteredTextToTestReqLetter = True
                                            'drdsTank("MODIFIED") = True
                                            'drdsFac("MODIFIED") = True
                                            'drdsOwner("MODIFIED") = True
                                            'End If
                                            'Else
                                            '           alTnkSecondaryDate.Add(CDate("1/1/1900"))

                                            '          drCalInfo = dtCalInfo.NewRow
                                            '         drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                            '        drCalInfo("NAME") = drdsFac("NAME")
                                            '       drCalInfo("CITY") = drdsFac("CITY")
                                            '      drCalInfo("MONTH") = dt.Month
                                            '     drCalInfo("INFO") = "Last inspection date of the tank secondary containment is unknown. This information is needed as soon a possible." + vbCrLf + _
                                            '                        "Please update my records to reflect that this test was accomplished on _______________."
                                            '   dtCalInfo.Rows.Add(drCalInfo)


                                            ' End If

                                            'End If




                                            'DateElectronicDeviceInspected
                                            If drdsTank("Tank_LD_Num") = 339 Then

                                                If Not drdsTank("DateElectronicDeviceInspected") Is DBNull.Value Then


                                                    dt = drdsTank("DateElectronicDeviceInspected")
                                                    dt = dt.Date
                                                    dt = DateAdd(DateInterval.Year, 1, dt)

                                                    If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then

                                                        If Not alTnkElectronicDate.Contains(dt.ToShortDateString) Then

                                                            alTnkElectronicDate.Add(dt.ToShortDateString)

                                                            drCalInfo = dtCalInfo.NewRow
                                                            drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                            drCalInfo("NAME") = drdsFac("NAME")
                                                            drCalInfo("CITY") = drdsFac("CITY")
                                                            drCalInfo("MONTH") = dt.Month
                                                            drCalInfo("INFO") = "Testing of tank electronic interstitial monitoring devices must be accomplished by " + dt.ToShortDateString + "." + vbCrLf + _
                                                                                "Please update my records to reflect that this test was accomplished on _______________."
                                                            dtCalInfo.Rows.Add(drCalInfo)

                                                        End If

                                                        bolEnteredTextToTestReqLetter = True
                                                        drdsTank("MODIFIED") = True
                                                        drdsFac("MODIFIED") = True
                                                        drdsOwner("MODIFIED") = True
                                                    End If

                                                Else
                                                    alTnkElectronicDate.Add(CDate("1/1/1900"))

                                                    drCalInfo = dtCalInfo.NewRow
                                                    drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                    drCalInfo("NAME") = drdsFac("NAME")
                                                    drCalInfo("CITY") = drdsFac("CITY")
                                                    drCalInfo("MONTH") = dt.Month
                                                    drCalInfo("INFO") = "Last testing date of tank electronic interstitial monitoring devices is unknown." + vbCrLf + _
                                                                        "Please update my records to reflect that this test was accomplished on _______________."
                                                    dtCalInfo.Rows.Add(drCalInfo)

                                                End If

                                            End If

                                            'DateATGLastInspected
                                            If drdsTank("Tank_LD_Num") = 336 Then



                                                If Not drdsTank("DateATGLastInspected") Is DBNull.Value Then

                                                    dt = drdsTank("DateATGLastInspected")
                                                    dt = dt.Date
                                                    dt = DateAdd(DateInterval.Year, 1, dt)

                                                    If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then

                                                        If Not alTnkATGDate.Contains(dt.ToShortDateString) Then

                                                            alTnkATGDate.Add(dt.ToShortDateString)

                                                            drCalInfo = dtCalInfo.NewRow
                                                            drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                            drCalInfo("NAME") = drdsFac("NAME")
                                                            drCalInfo("CITY") = drdsFac("CITY")
                                                            drCalInfo("MONTH") = dt.Month
                                                            drCalInfo("INFO") = "Inspection of automatic tank gauging equipment must be accomplished by " + dt.ToShortDateString + "." + vbCrLf + _
                                                                                "Please update my records to reflect that this test was accomplished on _______________."
                                                            dtCalInfo.Rows.Add(drCalInfo)

                                                        End If

                                                        bolEnteredTextToTestReqLetter = True
                                                        drdsTank("MODIFIED") = True
                                                        drdsFac("MODIFIED") = True
                                                        drdsOwner("MODIFIED") = True
                                                    End If

                                                Else
                                                    alTnkATGDate.Add(CDate("1/1/1900"))

                                                    drCalInfo = dtCalInfo.NewRow
                                                    drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                    drCalInfo("NAME") = drdsFac("NAME")
                                                    drCalInfo("CITY") = drdsFac("CITY")
                                                    drCalInfo("MONTH") = dt.Month
                                                    drCalInfo("INFO") = "Last inspection date of automatic tank gauging equipment is unknown. " + vbCrLf + _
                                                                        "Please update my records to reflect that this test was accomplished on _______________."
                                                    dtCalInfo.Rows.Add(drCalInfo)

                                                End If

                                            End If


                                        End If


                                        ' LAST TCP DATE
                                        If Not drdsTank("TANKMODDESC") Is DBNull.Value Then


                                            If drdsTank("TANKMODDESC").ToString.IndexOf("Cathodically Protected") > -1 Then


                                                If Not drdsTank("CP DATE") Is DBNull.Value Then

                                                    dt = drdsTank("CP DATE")
                                                    dt = dt.Date
                                                    dt = DateAdd(DateInterval.Year, 3, dt)

                                                    If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then

                                                        If Not alTnkLastTCPDate.Contains(dt.ToShortDateString) Then

                                                            alTnkLastTCPDate.Add(dt.ToShortDateString)


                                                            drCalInfo = dtCalInfo.NewRow
                                                            drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                            drCalInfo("NAME") = drdsFac("NAME")
                                                            drCalInfo("CITY") = drdsFac("CITY")
                                                            drCalInfo("MONTH") = dt.Month
                                                            drCalInfo("INFO") = "Testing of the tank cathodic protection must be accomplished by " + dt.ToShortDateString + "." + vbCrLf + _
                                                                                "Please update my records to reflect that this test was accomplished on _______________."
                                                            dtCalInfo.Rows.Add(drCalInfo)

                                                        End If

                                                        bolEnteredTextToTestReqLetter = True
                                                        drdsTank("MODIFIED") = True
                                                        drdsFac("MODIFIED") = True
                                                        drdsOwner("MODIFIED") = True
                                                    End If
                                                Else
                                                    alTnkLastTCPDate.Add(CDate("1/1/1900"))

                                                    drCalInfo = dtCalInfo.NewRow
                                                    drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                    drCalInfo("NAME") = drdsFac("NAME")
                                                    drCalInfo("CITY") = drdsFac("CITY")
                                                    drCalInfo("MONTH") = dt.Month
                                                    drCalInfo("INFO") = "Last testing date of the tank cathodic protection is unknown." + vbCrLf + _
                                                                        "Please update my records to reflect that this test was accomplished on _______________."
                                                    dtCalInfo.Rows.Add(drCalInfo)


                                                End If

                                            End If

                                        End If

                                        ' LINED DUE
                                        If Not drdsTank("TANKMODDESC") Is DBNull.Value Then

                                            If drdsTank("TANKMODDESC").ToString.IndexOf("Lined Interior") > -1 Then

                                                'enableTankLIInspectedDate = True
                                                dt = IIf(drdsTank("LI INSTALL") Is DBNull.Value, CDate("01/01/0001"), drdsTank("LI INSTALL"))
                                                dt1 = IIf(drdsTank("LI INSPECTED") Is DBNull.Value, CDate("01/01/0001"), drdsTank("LI INSPECTED"))
                                                dt = dt.Date
                                                dt1 = dt1.Date


                                                If dt > CDate("01/01/0001") Then


                                                    dt = DateAdd(DateInterval.Year, 10, dt)
                                                    dt1 = DateAdd(DateInterval.Year, 5, dt1)

                                                    If dt1 >= dt Then
                                                        dt = dt1
                                                    End If



                                                End If


                                                If Date.Compare(dt, CDate("01/01/0001")) <> 0 AndAlso dt >= dtProcessingStart.AddMonths(-3) Then

                                                    If Not drdsTank("TANKMODDESC").ToString.IndexOf("Cathodically Protected/Lined Interior") > -1 Then

                                                        If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then

                                                            If Not alTnkLinedDate.Contains(dt.ToShortDateString) Then

                                                                alTnkLinedDate.Add(dt.ToShortDateString)


                                                                drCalInfo = dtCalInfo.NewRow
                                                                drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                drCalInfo("NAME") = drdsFac("NAME")
                                                                drCalInfo("CITY") = drdsFac("CITY")
                                                                drCalInfo("MONTH") = dt.Month
                                                                drCalInfo("INFO") = "Inspection of the tank interior lining must be accomplished by " + dt.ToShortDateString + "." + vbCrLf + _
                                                                                    "Please update my records to reflect that this test was accomplished on _______________."
                                                                dtCalInfo.Rows.Add(drCalInfo)

                                                            End If

                                                            bolEnteredTextToTestReqLetter = True
                                                            drdsTank("MODIFIED") = True
                                                            drdsFac("MODIFIED") = True
                                                            drdsOwner("MODIFIED") = True

                                                        End If

                                                    End If
                                                Else
                                                    If Not drdsTank("TANKMODDESC").ToString.IndexOf("Cathodically Protected/Lined Interior") > -1 Then


                                                        alTnkLinedDate.Add(CDate("1/1/1900"))

                                                        drCalInfo = dtCalInfo.NewRow
                                                        drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                        drCalInfo("NAME") = drdsFac("NAME")
                                                        drCalInfo("CITY") = drdsFac("CITY")
                                                        drCalInfo("MONTH") = dt.Month
                                                        drCalInfo("INFO") = "Last inspection date of the tank interior lining is out of date or unknown." + vbCrLf + _
                                                                            "Please update my records to reflect that this test was accomplished on _______________."
                                                        dtCalInfo.Rows.Add(drCalInfo)
                                                    End If


                                                End If

                                            End If
                                        End If


                                        ' IF CIU
                                        If drdsTank("STATUS").ToString.IndexOf("Currently In Use") > -1 Then


                                            ' TANK TT DUE / IC EXPIRES
                                            ' Show Tank TT Due only if IC Expires is false
                                            If Not drdsTank("TANKLD") Is DBNull.Value Then

                                                If drdsTank("TANKLD").ToString.IndexOf("Inventory Control/Precision Tightness Testing") > -1 Then

                                                    ' ICExpires
                                                    dt = IIf(drdsTank("INSTALLED") Is DBNull.Value, CDate("01/01/0001"), drdsTank("INSTALLED"))
                                                    dt1 = IIf(drdsTank("TCPINSTALLDATE") Is DBNull.Value, CDate("01/01/0001"), drdsTank("TCPINSTALLDATE"))
                                                    dt = dt.Date
                                                    dt1 = dt1.Date

                                                    If Date.Compare(dt, dt1) < 0 Then
                                                        dt = dt1
                                                    End If

                                                    dt1 = IIf(drdsTank("LI INSTALL") Is DBNull.Value, CDate("01/01/0001"), drdsTank("LI INSTALL"))

                                                    dt1 = dt1.Date

                                                    If Date.Compare(dt, dt1) < 0 Then
                                                        dt = dt1
                                                    End If

                                                    dt = DateAdd(DateInterval.Year, 10, dt)

                                                    If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then

                                                        If Not alTnkICExpiresDate.Contains(dt.ToShortDateString) Then

                                                            alTnkICExpiresDate.Add(dt.ToShortDateString)

                                                            drCalInfo = dtCalInfo.NewRow
                                                            drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                            drCalInfo("NAME") = drdsFac("NAME")
                                                            drCalInfo("CITY") = drdsFac("CITY")
                                                            drCalInfo("MONTH") = dt.Month
                                                            drCalInfo("INFO") = "Please be aware that Inventory Control/Precision Tightness Testing is only a valid method of tank leak " + _
                                                                                "detection for a period of 10 years following tank installation or upgrade. Therefore, you must choose another " + _
                                                                                "method of tank leak detection by no later than " + dt.ToShortDateString + "."
                                                            dtCalInfo.Rows.Add(drCalInfo)


                                                        End If

                                                        bolEnteredTextToTestReqLetter = True
                                                        ' no need to roll over date - according to stefanie
                                                    Else

                                                        ' TANK TT DUE
                                                        dt = IIf(drdsTank("TT DATE") Is DBNull.Value, CDate("01/01/0001"), drdsTank("TT DATE"))
                                                        dt1 = IIf(drdsTank("INSTALLED") Is DBNull.Value, CDate("01/01/0001"), drdsTank("INSTALLED"))
                                                        dt = dt.Date
                                                        dt1 = dt1.Date

                                                        If Date.Compare(dt, dt1) < 0 Then
                                                            dt = dt1
                                                        End If

                                                        dt1 = IIf(drdsTank("TCPINSTALLDATE") Is DBNull.Value, CDate("01/01/0001"), drdsTank("TCPINSTALLDATE"))
                                                        dt1 = dt1.Date

                                                        If Date.Compare(dt, dt1) < 0 Then
                                                            dt = dt1
                                                        End If

                                                        dt1 = IIf(drdsTank("LI INSTALL") Is DBNull.Value, CDate("01/01/0001"), drdsTank("LI INSTALL"))
                                                        dt1 = dt1.Date

                                                        If Date.Compare(dt, dt1) < 0 Then
                                                            dt = dt1
                                                        End If

                                                        dt = DateAdd(DateInterval.Year, 5, dt)
                                                        dt1 = CDate("12/22/1998")

                                                        If Date.Compare(dt, dt1) < 0 Then
                                                            dt = dt1
                                                        End If

                                                        If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then


                                                            If Not alTnkTTDate.Contains(dt.ToShortDateString) Then

                                                                alTnkTTDate.Add(dt.ToShortDateString)

                                                                drCalInfo = dtCalInfo.NewRow
                                                                drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                drCalInfo("NAME") = drdsFac("NAME")
                                                                drCalInfo("CITY") = drdsFac("CITY")
                                                                drCalInfo("MONTH") = dt.Month
                                                                drCalInfo("INFO") = "Precision tightness testing of the tanks must be accomplished by " + dt.ToShortDateString + "." + vbCrLf + _
                                                                                    "Please update my records to reflect that this test was accomplished on _______________."
                                                                dtCalInfo.Rows.Add(drCalInfo)
                                                            End If

                                                            bolEnteredTextToTestReqLetter = True
                                                            drdsTank("MODIFIED") = True
                                                            drdsFac("MODIFIED") = True
                                                            drdsOwner("MODIFIED") = True
                                                        End If

                                                    End If

                                                End If ' if tankld = inventory control/precision tightness testing
                                            End If ' if tankld is null

                                        End If ' if ciu

                                    End If ' status is ciu / tosi
                                End If ' status is null


                                'Pipe Loop for CAP monthly report
                                For Each drdsPipe In ds.Tables(3).Select("FACILITY_ID = " + drdsFac("FACILITY_ID").ToString + " AND [TANK ID] = " + drdsTank("TANK ID").ToString) ' pipe


                                    drPipe = dtPipe.NewRow
                                    drPipe("TANK ID") = drdsTank("TANK ID")
                                    drPipe("PIPE ID") = drdsPipe("PIPE ID")

                                    ' check pipe conditions
                                    ' if pipe modified, set MODIFIED column value to true
                                    If Not drdsPipe("STATUS") Is DBNull.Value Then

                                        If drdsPipe("STATUS").ToString.IndexOf("Currently In Use") > -1 Or drdsPipe("STATUS").ToString.IndexOf("Temporarily Out of Service Indefinitely") > -1 Then

                                            'DateSheerValueTest
                                            If drdsPipe("PipeType") = 266 AndAlso drdsPipe("STATUS").ToString.IndexOf("Currently In Use") > -1 AndAlso Not drdsTank("TANKEMERGEN").ToString = "True" Then

                                                If Not drdsPipe("DateSheerValueTest") Is DBNull.Value Then

                                                    dt = drdsPipe("DateSheerValueTest")
                                                    dt = dt.Date
                                                    dt = DateAdd(DateInterval.Year, 1, dt)

                                                    If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then

                                                        If Not alPipeSheerDate.Contains(dt.ToShortDateString) Then

                                                            alPipeSheerDate.Add(dt.ToShortDateString)

                                                            drCalInfo = dtCalInfo.NewRow
                                                            drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                            drCalInfo("NAME") = drdsFac("NAME")
                                                            drCalInfo("CITY") = drdsFac("CITY")
                                                            drCalInfo("MONTH") = dt.Month
                                                            drCalInfo("INFO") = "Testing of pressurized piping shear valves must be accomplished by " + dt.ToShortDateString + "." + vbCrLf + _
                                                                                "Please update my records to reflect that this test was accomplished on _______________."
                                                            dtCalInfo.Rows.Add(drCalInfo)

                                                        End If

                                                        bolEnteredTextToTestReqLetter = True
                                                        drdsPipe("MODIFIED") = True
                                                        drdsTank("MODIFIED") = True
                                                        drdsFac("MODIFIED") = True
                                                        drdsOwner("MODIFIED") = True
                                                    End If
                                                Else
                                                    alPipeSheerDate.Add(CDate("1/1/1900"))

                                                    drCalInfo = dtCalInfo.NewRow
                                                    drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                    drCalInfo("NAME") = drdsFac("NAME")
                                                    drCalInfo("CITY") = drdsFac("CITY")
                                                    drCalInfo("MONTH") = dt.Month
                                                    drCalInfo("INFO") = "Last testing date of pressurized piping shear valves is unknown." + vbCrLf + _
                                                                        "Please update my records to reflect that this test was accomplished on _______________."
                                                    dtCalInfo.Rows.Add(drCalInfo)

                                                End If
                                            End If


                                            'DateSecondaryContainmentInspect
                                            If drdsPipe("Pipe_LD_Num") = 242 And drdsPipe("STATUS").ToString.IndexOf("Currently In Use") > -1 Then

                                                If Not drdsPipe("DateSecondaryContainmentInspect") Is DBNull.Value Then

                                                    dt = drdsPipe("DateSecondaryContainmentInspect")
                                                    dt = dt.Date
                                                    dt = DateAdd(DateInterval.Year, 1, dt)

                                                    If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                        If Not alPipeSecondaryDate.Contains(dt.ToShortDateString) Then

                                                            alPipeSecondaryDate.Add(dt.ToShortDateString)

                                                            drCalInfo = dtCalInfo.NewRow
                                                            drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                            drCalInfo("NAME") = drdsFac("NAME")
                                                            drCalInfo("CITY") = drdsFac("CITY")
                                                            drCalInfo("MONTH") = dt.Month
                                                            drCalInfo("INFO") = "Inspection of the pipe secondary containment sumps must be accomplished by " + dt.ToShortDateString + "." + vbCrLf + _
                                                                                "Please update my records to reflect that this test was accomplished on _______________."
                                                            dtCalInfo.Rows.Add(drCalInfo)
                                                        End If

                                                        bolEnteredTextToTestReqLetter = True
                                                        drdsPipe("MODIFIED") = True
                                                        drdsTank("MODIFIED") = True
                                                        drdsFac("MODIFIED") = True
                                                        drdsOwner("MODIFIED") = True
                                                    End If
                                                Else
                                                    alPipeSecondaryDate.Add(CDate("1/1/1900"))

                                                    drCalInfo = dtCalInfo.NewRow
                                                    drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                    drCalInfo("NAME") = drdsFac("NAME")
                                                    drCalInfo("CITY") = drdsFac("CITY")
                                                    drCalInfo("MONTH") = dt.Month
                                                    drCalInfo("INFO") = "Last inspection date of the secondary containment sumps is unknown." + vbCrLf + _
                                                                        "Please update my records to reflect that this test was accomplished on _______________."
                                                    dtCalInfo.Rows.Add(drCalInfo)

                                                End If
                                            End If


                                            'DateElectronicDeviceInspect
                                            If drdsPipe("Pipe_LD_Num") = 243 And drdsPipe("STATUS").ToString.IndexOf("Currently In Use") > -1 Then


                                                If Not drdsPipe("DateElectronicDeviceInspect") Is DBNull.Value Then

                                                    dt = drdsPipe("DateElectronicDeviceInspect")
                                                    dt = dt.Date
                                                    dt = DateAdd(DateInterval.Year, 1, dt)

                                                    If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then

                                                        If Not alPipeElectronicDate.Contains(dt.ToShortDateString) Then

                                                            alPipeElectronicDate.Add(dt.ToShortDateString)

                                                            drCalInfo = dtCalInfo.NewRow
                                                            drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                            drCalInfo("NAME") = drdsFac("NAME")
                                                            drCalInfo("CITY") = drdsFac("CITY")
                                                            drCalInfo("MONTH") = dt.Month
                                                            drCalInfo("INFO") = "Testing of the line electronic interstitial monitoring devices must be accomplished by " + dt.ToShortDateString + "." + vbCrLf + _
                                                                                "Please update my records to reflect that this test was accomplished on _______________."
                                                            dtCalInfo.Rows.Add(drCalInfo)
                                                        End If

                                                        bolEnteredTextToTestReqLetter = True
                                                        drdsPipe("MODIFIED") = True
                                                        drdsTank("MODIFIED") = True
                                                        drdsFac("MODIFIED") = True
                                                        drdsOwner("MODIFIED") = True
                                                    End If
                                                Else
                                                    alPipeElectronicDate.Add(CDate("1/1/1900"))

                                                    drCalInfo = dtCalInfo.NewRow
                                                    drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                    drCalInfo("NAME") = drdsFac("NAME")
                                                    drCalInfo("CITY") = drdsFac("CITY")
                                                    drCalInfo("MONTH") = dt.Month
                                                    drCalInfo("INFO") = "Last testing date of the line electronic interstitial monitoring device is unknown." + vbCrLf + _
                                                                        "Please update my records to reflect that this test was accomplished on _______________."
                                                    dtCalInfo.Rows.Add(drCalInfo)

                                                End If
                                            End If

                                            ' PIPE CP DATE
                                            If Not drdsPipe("PIPE_MOD_DESC") Is DBNull.Value Then

                                                If drdsPipe("PIPE_MOD_DESC").ToString = "Cathodically Protected" Then

                                                    If Not drdsPipe("CP DATE") Is DBNull.Value Then

                                                        dt = drdsPipe("CP DATE")
                                                        dt = dt.Date
                                                        dt = DateAdd(DateInterval.Year, 3, dt)

                                                        If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then

                                                            If Not alPipeCPDate.Contains(dt.ToShortDateString) Then

                                                                alPipeCPDate.Add(dt.ToShortDateString)

                                                                drCalInfo = dtCalInfo.NewRow
                                                                drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                drCalInfo("NAME") = drdsFac("NAME")
                                                                drCalInfo("CITY") = drdsFac("CITY")
                                                                drCalInfo("MONTH") = dt.Month
                                                                drCalInfo("INFO") = "Testing of the pipe cathodic protection must be accomplished by " + dt.ToShortDateString + "." + vbCrLf + _
                                                                                    "Please update my records to reflect that this test was accomplished on _______________."
                                                                dtCalInfo.Rows.Add(drCalInfo)

                                                            End If

                                                            bolEnteredTextToTestReqLetter = True
                                                            drdsPipe("MODIFIED") = True
                                                            drdsTank("MODIFIED") = True
                                                            drdsFac("MODIFIED") = True
                                                            drdsOwner("MODIFIED") = True
                                                        End If
                                                    Else
                                                        alPipeCPDate.Add(CDate("1/1/1900"))

                                                        drCalInfo = dtCalInfo.NewRow
                                                        drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                        drCalInfo("NAME") = drdsFac("NAME")
                                                        drCalInfo("CITY") = drdsFac("CITY")
                                                        drCalInfo("MONTH") = dt.Month
                                                        drCalInfo("INFO") = "Last testing date of the pipe cathodic protection is unknown." + vbCrLf + _
                                                                            "Please update my records to reflect that this test was accomplished on _______________."
                                                        dtCalInfo.Rows.Add(drCalInfo)

                                                    End If
                                                End If
                                            End If

                                            ' PIPE TERM CP DATE
                                            Dim bolContinue As Boolean = False

                                            If Not drdsPipe("DISP CP TYPE") Is DBNull.Value Then

                                                If drdsPipe("DISP CP TYPE").ToString.IndexOf("Cathodically Protected") > -1 Then
                                                    If Not drdsPipe("TERM CP TEST") Is DBNull.Value Then
                                                        bolContinue = True
                                                    Else
                                                        alPipeTermCPDate.Add(CDate("1/1/1900"))

                                                        drCalInfo = dtCalInfo.NewRow
                                                        drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                        drCalInfo("NAME") = drdsFac("NAME")
                                                        drCalInfo("CITY") = drdsFac("CITY")
                                                        drCalInfo("MONTH") = dt.Month
                                                        drCalInfo("INFO") = "Last testing date of the piping flex connector cathodic protection is unknown." + vbCrLf + _
                                                                            "Please update my records to reflect that this test was accomplished on _______________."
                                                        dtCalInfo.Rows.Add(drCalInfo)

                                                    End If
                                                End If

                                            End If

                                            If Not bolContinue Then

                                                If Not drdsPipe("TANK CP TYPE") Is DBNull.Value Then
                                                    If drdsPipe("TANK CP TYPE").ToString.IndexOf("Cathodically Protected") > -1 Then
                                                        If Not drdsPipe("TERM CP TEST") Is DBNull.Value Then
                                                            bolContinue = True
                                                        Else
                                                            alPipeTermCPDate.Add(CDate("1/1/1900"))

                                                            drCalInfo = dtCalInfo.NewRow
                                                            drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                            drCalInfo("NAME") = drdsFac("NAME")
                                                            drCalInfo("CITY") = drdsFac("CITY")
                                                            drCalInfo("MONTH") = dt.Month
                                                            drCalInfo("INFO") = "Last testing date of the piping flex connector cathodic protection is unknown." + vbCrLf + _
                                                                                "Please update my records to reflect that this test was accomplished on _______________."
                                                            dtCalInfo.Rows.Add(drCalInfo)

                                                        End If
                                                    End If
                                                End If

                                            End If

                                            If bolContinue Then

                                                dt = drdsPipe("TERM CP TEST")
                                                dt = dt.Date
                                                dt = DateAdd(DateInterval.Year, 3, dt)

                                                If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then

                                                    If Not alPipeTermCPDate.Contains(dt.ToShortDateString) Then

                                                        alPipeTermCPDate.Add(dt.ToShortDateString)

                                                        drCalInfo = dtCalInfo.NewRow
                                                        drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                        drCalInfo("NAME") = drdsFac("NAME")
                                                        drCalInfo("CITY") = drdsFac("CITY")
                                                        drCalInfo("MONTH") = dt.Month
                                                        drCalInfo("INFO") = "Testing of the piping flex connector cathodic protection must be accomplished by " + dt.ToShortDateString + "." + vbCrLf + _
                                                                            "Please update my records to reflect that this test was accomplished on _______________."
                                                        dtCalInfo.Rows.Add(drCalInfo)

                                                    End If

                                                    bolEnteredTextToTestReqLetter = True
                                                    drdsPipe("MODIFIED") = True
                                                    drdsTank("MODIFIED") = True
                                                    drdsFac("MODIFIED") = True
                                                    drdsOwner("MODIFIED") = True
                                                End If
                                            End If

                                            ' IF CIU
                                            If drdsPipe("STATUS").ToString.IndexOf("Currently In Use") > -1 Then

                                                ' ALLD TEST DATE
                                                If Not drdsPipe("ALLD_TEST") Is DBNull.Value Then

                                                    ' If drdsPipe("ALLD_TEST").ToString = "Mechanical" Then

                                                    If drdsPipe("PIPE_TYPE_DESC").ToString = "Pressurized" And (Not (drdsPipe("PIPE_LD").ToString = "Deferred")) Then

                                                        If Not drdsPipe("ALLD_TEST_DATE") Is DBNull.Value Then

                                                            dt = drdsPipe("ALLD_TEST_DATE")
                                                            dt = dt.Date
                                                            dt = DateAdd(DateInterval.Year, 1, dt)

                                                            If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then

                                                                If Not alPipeALLDDate.Contains(dt.ToShortDateString) Then
                                                                    alPipeALLDDate.Add(dt.ToShortDateString)

                                                                    drCalInfo = dtCalInfo.NewRow
                                                                    drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                    drCalInfo("NAME") = drdsFac("NAME")
                                                                    drCalInfo("CITY") = drdsFac("CITY")
                                                                    drCalInfo("MONTH") = dt.Month
                                                                    drCalInfo("INFO") = "Testing of the automatic line leak detector must be accomplished by " + dt.ToShortDateString + "." + vbCrLf + _
                                                                                        "Please update my records to reflect that this test was accomplished on _______________."
                                                                    dtCalInfo.Rows.Add(drCalInfo)

                                                                End If

                                                                bolEnteredTextToTestReqLetter = True
                                                                drdsPipe("MODIFIED") = True
                                                                drdsTank("MODIFIED") = True
                                                                drdsFac("MODIFIED") = True
                                                                drdsOwner("MODIFIED") = True
                                                            End If
                                                        Else
                                                            alPipeALLDDate.Add(CDate("1/1/1900"))

                                                            drCalInfo = dtCalInfo.NewRow
                                                            drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                            drCalInfo("NAME") = drdsFac("NAME")
                                                            drCalInfo("CITY") = drdsFac("CITY")
                                                            drCalInfo("MONTH") = dt.Month
                                                            drCalInfo("INFO") = "Last testing date of the automatic line leak detector is unknown." + vbCrLf + _
                                                                                "Please update my records to reflect that this test was accomplished on _______________."
                                                            dtCalInfo.Rows.Add(drCalInfo)


                                                        End If

                                                    End If
                                                End If


                                                ' PIPE LINE
                                                If Not drdsPipe("PIPE_LD") Is DBNull.Value And drdsPipe("STATUS").ToString.IndexOf("Currently In Use") > -1 Then

                                                    If drdsPipe("PIPE_LD").ToString = "Line Tightness Testing" Then

                                                        If Not drdsPipe("PIPE_TYPE_DESC") Is DBNull.Value Then



                                                            If Not drdsPipe("TT DATE") Is DBNull.Value Then


                                                                ' US SUCTION
                                                                If drdsPipe("PIPE_TYPE_DESC").ToString = "U.S. Suction" Then


                                                                    dt = drdsPipe("TT DATE")
                                                                    dt = dt.Date
                                                                    dt = DateAdd(DateInterval.Year, 3, dt)

                                                                    If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then

                                                                        If Not alPipeLineUSDate.Contains(dt.ToShortDateString) Then
                                                                            alPipeLineUSDate.Add(dt.ToShortDateString)

                                                                            drCalInfo = dtCalInfo.NewRow
                                                                            drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                            drCalInfo("NAME") = drdsFac("NAME")
                                                                            drCalInfo("CITY") = drdsFac("CITY")
                                                                            drCalInfo("MONTH") = dt.Month
                                                                            drCalInfo("INFO") = "Precision tightness testing of the 'U.S.' suction piping must be accomplished by " + dt.ToShortDateString + "." + vbCrLf + _
                                                                                                "Please update my records to reflect that this test was accomplished on _______________."
                                                                            dtCalInfo.Rows.Add(drCalInfo)

                                                                        End If

                                                                        bolEnteredTextToTestReqLetter = True
                                                                        drdsPipe("MODIFIED") = True
                                                                        drdsTank("MODIFIED") = True
                                                                        drdsFac("MODIFIED") = True
                                                                        drdsOwner("MODIFIED") = True
                                                                    End If


                                                                    ' PRESSURIZED
                                                                ElseIf drdsPipe("PIPE_TYPE_DESC").ToString = "Pressurized" And drdsPipe("STATUS").ToString.IndexOf("Currently In Use") > -1 Then

                                                                    dt = drdsPipe("TT DATE")
                                                                    dt = dt.Date
                                                                    dt = DateAdd(DateInterval.Year, 1, dt)

                                                                    If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then

                                                                        If Not alPipeLinePressDate.Contains(dt.ToShortDateString) Then

                                                                            alPipeLinePressDate.Add(dt.ToShortDateString)

                                                                            drCalInfo = dtCalInfo.NewRow
                                                                            drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                            drCalInfo("NAME") = drdsFac("NAME")
                                                                            drCalInfo("CITY") = drdsFac("CITY")
                                                                            drCalInfo("MONTH") = dt.Month
                                                                            drCalInfo("INFO") = "Precision tightness testing of the pressurized piping must be accomplished by " + dt.ToShortDateString + "." + vbCrLf + _
                                                                                                "Please update my records to reflect that this test was accomplished on _______________."
                                                                            dtCalInfo.Rows.Add(drCalInfo)
                                                                        End If

                                                                        bolEnteredTextToTestReqLetter = True
                                                                        drdsPipe("MODIFIED") = True
                                                                        drdsTank("MODIFIED") = True
                                                                        drdsFac("MODIFIED") = True
                                                                        drdsOwner("MODIFIED") = True
                                                                    End If

                                                                End If

                                                            Else
                                                                ' US SUCTION
                                                                If drdsPipe("PIPE_TYPE_DESC").ToString = "U.S. Suction" Then


                                                                    alPipeLineUSDate.Add(CDate("1/1/1900"))

                                                                    drCalInfo = dtCalInfo.NewRow
                                                                    drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                    drCalInfo("NAME") = drdsFac("NAME")
                                                                    drCalInfo("CITY") = drdsFac("CITY")
                                                                    drCalInfo("MONTH") = dt.Month
                                                                    drCalInfo("INFO") = "Last test date of precision tightness of the 'U.S.' suction piping is unknown." + vbCrLf + _
                                                                                                "Please update my records to reflect that this test was accomplished on _______________."
                                                                    dtCalInfo.Rows.Add(drCalInfo)



                                                                    ' PRESSURIZED
                                                                ElseIf drdsPipe("PIPE_TYPE_DESC").ToString = "Pressurized" And drdsPipe("STATUS").ToString.IndexOf("Currently In Use") > -1 Then





                                                                    alPipeLinePressDate.Add(CDate("1/1/1900"))

                                                                    drCalInfo = dtCalInfo.NewRow
                                                                    drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                    drCalInfo("NAME") = drdsFac("NAME")
                                                                    drCalInfo("CITY") = drdsFac("CITY")
                                                                    drCalInfo("MONTH") = dt.Month
                                                                    drCalInfo("INFO") = "Last test date of precision tightness of the pressurized piping is unknown." + vbCrLf + _
                                                                                        "Please update my records to reflect that this test was accomplished on _______________."
                                                                    dtCalInfo.Rows.Add(drCalInfo)
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If ' if ciu

                                        End If ' if status is ciu / tosi

                                    End If ' status is null

                                    If drdsPipe("MODIFIED") = True Then
                                        dtPipe.Rows.Add(drPipe)
                                    End If

                                Next ' pipe

                                If drdsTank("MODIFIED") = True Then
                                    dtTank.Rows.Add(drTank)
                                End If

                            Next ' tank

                            If drdsFac("MODIFIED") = True Then
                                dtFac.Rows.Add(drFac)

                            End If

                        Next ' facility

                        ' create assistance letter for owner
                        If bolAddedSectionForOwner Then

                            'AddCAPMonthlyComplianceAssistanceLetter(drdsOwner, docAssist, bolEnteredTextToAssistLetter, strAssistTemplate)

                            bolEnteredTextToAssistLetter = True
                            'Add Owner IDs for printing labels

                            If strOwnerIDs = String.Empty Then
                                strOwnerIDs = drdsOwner("OWNER_ID").ToString.Trim
                            Else
                                strOwnerIDs += "," + drdsOwner("OWNER_ID").ToString.Trim
                            End If

                        End If

                        'End With ' With docAssist

                        If drdsOwner("MODIFIED") = True Then
                            dtOwner.Rows.Add(drOwner)
                        End If
                    End If



                    If dtCalInfo.Rows.Count > 0 AndAlso bolEnteredTextToTestReqLetter Then


                        Try
                            pOwn.ClearCAPAnnualCalendar(processingYear, drdsOwner("OWNER_ID"), 1, processingMonth)

                            For Each drCal As DataRow In dtCalInfo.Rows

                                Dim facID As Integer = drCal("FACILITY_ID")
                                Dim OwnerID As Integer = drdsOwner("OWNER_ID")
                                Dim OwnerNameStr = drdsOwner("OWNERNAME").ToString.ToUpper
                                Dim City = drCal("CITY").ToString
                                Dim facility = drCal("NAME").ToString
                                Dim requirements = drCal("INFO").ToString
                                Dim month As Integer = drCal("MONTH")

                                Try
                                    pOwn.SaveCAPAnnualCalendar(processingYear, _
                                            OwnerID, _
                                            OwnerNameStr, _
                                            month, _
                                            facID, _
                                            facility, _
                                            City, _
                                            requirements, _
                                             _container.AppUser.ID, 1, processingMonth)
                                Catch ex As Exception
                                    If ex.ToString.IndexOf("PK_") = -1 Then
                                        Throw ex
                                    End If
                                End Try

                            Next
                        Catch ex As Exception
                            Throw ex
                        End Try

                    End If

                Next ' owner

                Try
                    _container.ActiveForm.Text = oldText
                Catch
                End Try



                    'Insert all the Owner details for printing Labels
                    pOwn.RunSQLQuery("EXEC spPutCAPLabels '" + strOwnerIDs + "','" + processingMonthYear + "','Monthly'")

                    Return True

                '      End With ' with wordapp
            Else

                MsgBox("No Records Found")

                ' delete the docs created at the top as no text was entered to the files
                bolDeleteFilesCreated = True

                Return False
            End If


        Catch ex As Exception
            bolDeleteFilesCreated = True
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()

        Finally
            _container.Text = oldText
        End Try
    End Function

    Private Function GetCapMonthlyNoticeOfTestReqHeading(ByVal wordApp As Word.Application) As String
        Dim doc As Word.Document
        Dim strReturnValue As String = ""
        Try
            doc = wordApp.Documents.Open(TmpltPath + "CAP\CapMonthlyTestingReqHeading.doc", , True, , , , , , , , , False)
            If Not doc Is Nothing Then
                doc.Activate()
                If doc.Tables.Count > 0 Then
                    strReturnValue = doc.Tables.Item(1).Cell(1, 1).Range.Text
                End If
            End If
        Catch ex As Exception
            Throw ex
        Finally
            If Not doc Is Nothing Then
                doc.Close(False)
            End If
        End Try
        Return strReturnValue
    End Function
    Private Function AddCapMontlyNoticeOfTestReqFacilityDetails(ByVal drFac As DataRow, ByRef doc As Word.Document, ByVal points As Integer, ByVal headingText As String, ByVal ownerName As String) As Integer
        Try
            ' insert Facility Details
            With doc
                doc.Activate()

                points = Me.AddCapMonhlyPageBreak(points + 10, points, doc, headingText, ownerName)

                oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                oPara.Range.Font.Name = "Arial"
                oPara.Range.Font.Size = 10
                oPara.Range.Font.Bold = 1
                Dim str As String = "FAC. I.D. #" + drFac("FACILITY_ID").ToString + " " + drFac("NAME").ToString.Trim + ", " + drFac("ADDRESS_LINE_ONE").ToString.Trim + ", " + drFac("CITY").ToString.Trim
                oPara.Range.Text = str
                oPara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                oPara.Range.InsertParagraphAfter()

                InsertLines(1, doc)

                Return points + 2
            End With

        Catch ex As Exception
            Throw ex
        End Try
    End Function


    Private Function AddCapMonhlyPageBreak(ByVal checkPoint As Integer, ByVal currentPoint As Integer, ByVal doc As Word.Document, ByVal headingText As String, ByVal ownerName As String) As Integer

        If checkPoint > 44 Then

            Try
                With doc

                    doc.Activate()


                    WordApp.Selection.EndKey(Unit:=Word.WdUnits.wdStory)
                    doc.Application.Selection.InsertBreak(Word.WdBreakType.wdPageBreak)


                    ' insert Blank
                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                    'oPara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                    oPara.Range.Font.Name = "Arial"
                    oPara.Range.Font.Size = 10
                    oPara.Range.Font.Bold = 0
                    oPara.Range.Text = String.Format("{0}{0}{0}{0}", vbCrLf)
                    oPara.Range.InsertParagraphAfter()
                    'InsertLines(1, doc)


                    ' insert heading 
                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                    'oPara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                    oPara.Range.Font.Name = "Arial"
                    oPara.Range.Font.Size = 10
                    oPara.Range.Font.Bold = 0
                    oPara.Range.Text = headingText
                    oPara.Range.InsertParagraphAfter()
                    'InsertLines(1, doc)

                    ' insert owner name
                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                    oPara.Range.Font.Name = "Arial"
                    oPara.Range.Font.Size = 11
                    oPara.Range.Font.Bold = 1
                    oPara.Range.Text = ownerName
                    oPara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    oPara.Range.InsertParagraphAfter()
                    InsertLines(1, doc)
                End With
            Catch ex As Exception
                Throw ex
            End Try

            currentPoint = 0

        End If

        Return currentPoint

    End Function

    Private Sub AddCAPMonthlyNoticeOfTestReqHeading(ByVal headingText As String, ByVal ownerName As String, ByRef doc As Word.Document, ByVal bolAddSectionBreak As Boolean)
        Try
            With doc
                doc.Activate()

                If bolAddSectionBreak Then
                    ' insert section break
                    WordApp.Selection.EndKey(Unit:=Word.WdUnits.wdStory)
                    doc.Application.Selection.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage)
                End If

                ' insert heading 
                oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                'oPara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                oPara.Range.Font.Name = "Arial"
                oPara.Range.Font.Size = 10
                oPara.Range.Font.Bold = 0
                oPara.Range.Text = headingText
                oPara.Range.InsertParagraphAfter()
                'InsertLines(1, doc)

                ' insert owner name
                oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                oPara.Range.Font.Name = "Arial"
                oPara.Range.Font.Size = 11
                oPara.Range.Font.Bold = 1
                oPara.Range.Text = ownerName
                oPara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                oPara.Range.InsertParagraphAfter()
                InsertLines(1, doc)
            End With
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub AddCAPMonthlyComplianceAssistanceLetter(ByVal drOwner As DataRow, ByRef doc As Word.Document, ByVal bolAddSectionBreak As Boolean, ByVal strAssistTemplate As String)
        Dim colParams As New Specialized.NameValueCollection
        Dim strKey As String = String.Empty
        Dim strValue As String = String.Empty

        Try
            With doc
                doc.Activate()
                If bolAddSectionBreak Then
                    ' insert section break
                    WordApp.Selection.EndKey(Unit:=Word.WdUnits.wdStory)
                    doc.Application.Selection.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage)

                    ' insert file
                    doc.Application.Selection.InsertFile(FILENAME:=strAssistTemplate, ConfirmConversions:=False, Link:=False, Attachment:=False)
                End If

                ' Build NameValueCollection with Tags and Values.
                colParams.Add("<Title>", "Signature Needed Letter")
                colParams.Add("<Date>", Format(Now, "MMMM d, yyyy"))

                colParams.Add("<Owner Name>", drOwner("OWNERNAME").ToString.Trim)
                colParams.Add("<Owner Address 1>", drOwner("ADDRESS_LINE_ONE").ToString.Trim)
                If drOwner("ADDRESS_TWO") Is DBNull.Value Then
                    colParams.Add("<Owner Address 2>", drOwner("CITY").ToString.Trim + ", " + drOwner("STATE").ToString.Trim + " " + drOwner("ZIP").ToString.Trim)
                    colParams.Add("<Owner City/State/Zip>", "")
                ElseIf drOwner("ADDRESS_TWO").ToString.Trim = String.Empty Then
                    colParams.Add("<Owner Address 2>", drOwner("CITY").ToString.Trim + ", " + drOwner("STATE").ToString.Trim + " " + drOwner("ZIP").ToString.Trim)
                    colParams.Add("<Owner City/State/Zip>", "")
                Else
                    colParams.Add("<Owner Address 2>", drOwner("ADDRESS_TWO").ToString.Trim)
                    colParams.Add("<Owner City/State/Zip>", drOwner("CITY").ToString.Trim + ", " + drOwner("STATE").ToString.Trim + " " + drOwner("ZIP").ToString.Trim)
                End If

                If drOwner("ORGANIZATION_ID") Is DBNull.Value Then
                    colParams.Add("<Owner Greeting>", drOwner("OWNERNAME").ToString.Trim + ":")
                ElseIf drOwner("ORGANIZATION_ID") = 0 Then
                    colParams.Add("<Owner Greeting>", drOwner("OWNERNAME").ToString.Trim + ":")
                Else
                    colParams.Add("<Owner Greeting>", "Dear " + drOwner("OWNERNAME").ToString.Trim + ":")
                End If

                Dim userInfoLocal As MUSTER.Info.UserInfo
                userInfoLocal = MusterContainer.AppUser.RetrieveCAEHead()
                colParams.Add("<User Phone>", CType(userInfoLocal.PhoneNumber, String))
                colParams.Add("<User>", userInfoLocal.Name)

                ' Find and Replace the TAGs with Values.
                For i As Integer = 0 To colParams.Count - 1
                    strKey = colParams.Keys(i).ToString
                    strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                    .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)
                Next
            End With
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function NewDateAdd(ByVal dt As DateTime, ByVal num As Integer, ByVal year As Integer) As DateTime
        '  If Date.Compare(dt, New Date(year - 1, 11, 1)) <= 0 Then
        Return DateAdd(DateInterval.Year, num, dt)
        ' Else
        '    Return dt
        '  End If

    End Function


    Sub SetupSystemToGenerateCAPYearly(ByVal mode As CapAnnualMode, ByVal prompt As Boolean, ByVal strOwnerName As String, ByVal fac_id As Integer, Optional ByVal year As Integer = 0)

        Dim owner As New BusinessLogic.pOwner

        Try


            Dim pass As Boolean = False
            Dim facs As String = String.Empty

            'check rights for Tank & Pipe
            If Not _container.AppUser.HasAccess(CType(UIUtilsGen.ModuleID.CAPProcess, Integer), MusterContainer.AppUser.UserKey, UIUtilsGen.EntityTypes.Tank) Then
                MessageBox.Show("You do not have Rights to Yearly CAP Processing")
                Exit Sub
            ElseIf Not _container.AppUser.HasAccess(CType(UIUtilsGen.ModuleID.CAPProcess, Integer), MusterContainer.AppUser.UserKey, UIUtilsGen.EntityTypes.Pipe) Then
                MessageBox.Show("You do not have Rights to Yearly CAP Processing")
                Exit Sub
            End If

            Dim strYear As String
            Dim ownerName As String = strOwnerName


            If year = 0 Then
                strYear = InputBox("Enter Year (XXXX) for Annual", , Today.Year.ToString)
                If strYear = String.Empty Then
                    Exit Sub
                ElseIf Not IsNumeric(strYear) Then
                    MsgBox("Invalid entry")
                    Exit Sub
                ElseIf strYear.Length < 4 Then
                    MsgBox("Please enter 4 digit Year")
                    SetupSystemToGenerateCAPYearly(mode, prompt, ownerName, -1, year)
                    Exit Sub
                End If
            Else
                strYear = String.Format("{0}", year)
            End If


            If prompt Then
                ownerName = InputBox("Enter Owner Name or Facility # (Leave Blank for Full Annual CAP Report)", , String.Empty)
            End If

            If ownerName = String.Empty Then
                ownerName = ""
            End If

            _container.Cursor = Cursors.WaitCursor


            pass = GenerateCAPYearly(mode, strYear, owner, ownerName, facs)


            If pass Then

                Dim frmReport As ReportDisplay

                frmReport = New ReportDisplay
                frmReport.MdiParent = _container

                Dim ownID As Integer = 0

                If Not _container.pOwn Is Nothing Then
                    ownID = _container.pOwn.ID
                End If

                Try

                    If prompt Then
                        strOwnerName = ownerName


                    End If
                    'If Not strOwnerName Is Nothing Then
                    '    If strOwnerName.Length = 0 OrElse facs <> String.Empty Then
                    '        strOwnerName = " "
                    '    End If
                    'End If



                    'frmReport.Show()
                    'If mode = CapAnnualMode.StaticByYear Then
                    '    frmReport.GenerateReport("CAP Annual Summary Report", New Object() {strYear, IIf(facs = String.Empty, DBNull.Value, facs), IIf(ownID = 0, DBNull.Value, ownID), strOwnerName}, String.Format("Annual CAP Summary Completed for {0}{1}. Would you like to create a PDF document of it?", IIf(ownerName = String.Empty, "all", IIf(ownerName = fac_id.ToString, "", ownerName)), IIf(fac_id <= 0, "", fac_id)))

                    'Else
                    '    frmReport.GenerateReport("Current CAP Summary Per Owner/Facility", New Object() {strYear, IIf(facs = String.Empty, DBNull.Value, facs), IIf(ownID <= 0, DBNull.Value, ownID), strOwnerName}, IIf(year = 0, String.Empty, "Current CAP Summary Completed. Would you like to create a PDF document of it?"))
                    'End If

                Catch ex As Exception
                    Dim MyErr As ErrorReport
                    MyErr = New ErrorReport(New Exception("Error in loading reports form : " & vbCrLf & ex.Message, ex))
                    MyErr.ShowDialog()
                End Try
            End If


        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            _container.Cursor = Cursors.Default

            owner = Nothing
        End Try

    End Sub

    Friend Function GenerateCAPYearly(ByVal mode As CapAnnualMode, ByVal processingYear As Integer, ByVal pOwn As MUSTER.BusinessLogic.pOwner, Optional ByVal OwnerName As String = "", Optional ByRef facs As String = "") As Boolean

        Dim ds As DataSet
        Dim oldText As String = "Process CAP Report - "

        If Not _container Is Nothing AndAlso Not _container.ActiveForm Is Nothing Then
            oldText = _container.ActiveForm.Text()
        End If


        Dim dsRelOwnerFac, dsRelFacTank, dsRelTankPipe As DataRelation
        Dim dtSummary As DataTable
        Dim nSummaryLineCount As Integer = 0
        Dim strSummaryDesc As String = String.Empty
        Dim dtCalInfo As DataTable
        Dim drSummary, drFac, drCalInfo As DataRow ' using different variables for less confusion

        Dim drdsOwner, drdsFac, drdsTank, drdsPipe As DataRow ' for looping through the tables
        Dim drdsFacs() As DataRow
        Dim ownID As Integer = 0
        Dim dtProcessingStart, dtProcessingEnd As Date
        'Dim strSummaryDocName As String = ""

        'Dim strCalDocName As String = ""
        'Dim strAssessCalDocName As String = ""
        'Dim strAssessNoCalDocName As String = ""
        'Dim strTemplate As String = ""
        'Dim strCalTemplate As String = ""
        'Dim strAssessCalTemplate As String = ""
        'Dim strAssessNoCalTemplate As String = ""
        Dim strUser As String = ""
        Dim strUserPhone As String = ""

        Dim bolDeleteFilesCreated As Boolean = False
        'Dim bolDeletedSummary As Boolean = False
        Dim bolDeletedAssessCal As Boolean = False
        Dim bolDeletedAssessNoCal As Boolean = False
        Dim bolDeletedCal As Boolean = False

        Dim headingText As String = ""
        Dim dt, dt1 As Date

        Dim bolEnteredTextToSummary As Boolean = False
        Dim bolEnteredTextToCalendar As Boolean = False
        Dim bolEnteredTextToAssessNoCal As Boolean = False
        Dim bolEnteredTextToAssessCal As Boolean = False

        Dim bolAddedSectionForOwner As Boolean = False
        Dim bolAddedSectionForFac As Boolean = False
        Dim bolAddedSummaryPeriodicTesting As Boolean = False
        Dim bolAddSummaryRoutineTesting As Boolean = False
        Dim bolOwnerNeedsCal As Boolean = False

        'Dim docSummary, docCal, docAssessCal, docAssessNoCal As Word.Document
        ' Dim docCal, docAssessCal, docAssessNoCal, docAssessCalDS, docAssessNoCalDS As Word.Document

        Dim alGW As New ArrayList
        Dim alATG As New ArrayList
        Dim alICTT As New ArrayList
        Dim alSIR As New ArrayList
        Dim alMTG As New ArrayList
        Dim alPipeTT As New ArrayList
        Dim alEALLD1 As New ArrayList
        Dim alEALLD2 As New ArrayList
        Dim alMALLD2 As New ArrayList
        Dim alElectInter As New ArrayList
        Dim alVisInter As New ArrayList
        Dim alImpress As New ArrayList
        Dim alSO As New ArrayList
        Dim alTOSI As New ArrayList

        Dim alTnkLastTCPDate As New ArrayList
        Dim alTnkLinedDate As New ArrayList
        Dim alTnkTTDate As New ArrayList
        Dim alTnkICExpiresDate As New ArrayList

        Dim alPipeCPDate As New ArrayList
        Dim alPipeALLDDate As New ArrayList
        Dim alPipeTermCPDate As New ArrayList
        Dim alPipeLineUSDate As New ArrayList
        Dim alPipeLinePressDate As New ArrayList

        Dim slMonth As New SortedList

        ' var to check if summary doc can be split. set to false if owner has any entry in current summary doc
        'Dim bolCanSplitSummary As Boolean = True

        'Dim strTime As String = String.Empty
        'Dim dtStartAll As DateTime = DateTime.Now
        'Dim dtEndAll As DateTime
        'Dim dtStart As DateTime = DateTime.Now
        'Dim dtEnd As DateTime
        'Dim ts As New TimeSpan

        Try
            If DOC_PATH = "\" Then
                MsgBox("Document Path Unspecified. Please give the path before generating the letter.")
                Exit Function
            End If

            If processingYear = 0 Then
                processingYear = Today.Year
            End If

            '' dtProcessingstart = Jan 1, processingYear
            dtProcessingStart = CDate("1/1/" + processingYear.ToString)
            '' dtProcessingEnd = Dec 31, processingYear
            dtProcessingEnd = CDate("12/31/" + processingYear.ToString)

            '''To avoid duplicate creation of Letters.
            ''strSummaryDocName = "REG_CAP_ANNUAL_SUMMARY_" + processingYear.ToString + "_0.doc"

            'strCalDocName = "REG_CAP_ANNUAL_CALENDAR_" + processingYear.ToString + OwnerName.Replace(" ", "") + ".doc"
            'strAssessCalDocName = "REG_CAP_ANNUAL_Assessment_Letter_With_Calendar_" + processingYear.ToString + OwnerName.Replace(" ", "") + ".doc"
            'strAssessNoCalDocName = "REG_CAP_ANNUAL_Assessment_Letter_No_Calendar_" + processingYear.ToString + OwnerName.Replace(" ", "") + ".doc"

            ''If FileExists(DOC_PATH + strSummaryDocName) Then
            ''    MsgBox("CAP Annual Summary for  " + processingYear.ToString + "  has been created already.")
            ''    Exit Sub
            ''End If
            ''If FileExists(DOC_PATH + strSummaryDocName.Substring(0, strSummaryDocName.Length - 5) + "1.doc") Then
            ''    MsgBox("CAP Annual Summary for  " + processingYear.ToString + "  has been created already.")
            ''    Exit Sub
            ''End If
            Dim fac As String = String.Empty

            If OwnerName.IndexOf(",") > -1 Then
                For Each str As String In OwnerName.Split(",")
                    If IsNumeric(str) Then
                        fac = String.Format("{0}{1}{2}", fac, IIf(fac.Length > 0, ",", String.Empty), str)
                    Else
                        fac = String.Empty
                        Exit For
                    End If
                Next
            ElseIf IsNumeric(OwnerName) Then
                fac = String.Format("{0}", OwnerName)
            End If

            facs = fac



            If 1 = 2 Then
                ''FileExists(DOC_PATH + strCalDocName) OrElse _
                ''FileExists(DOC_PATH + strAssessCalDocName) OrElse _
                ''FileExists(DOC_PATH + strAssessNoCalDocName) Then

                If MsgBox("Annual CAP documents for " + processingYear.ToString + IIf(OwnerName.Length > 0, String.Format(" -  for owner '{0}'", OwnerName), "") + "  has been created already.  Would you like to rebuild the CAP Annual report? ", MsgBoxStyle.YesNo) = MsgBoxResult.No Then

                    Exit Function
                Else
                    ''File.Delete(DOC_PATH + strCalDocName)
                    ''File.Delete(DOC_PATH + strAssessCalDocName)
                    ''File.Delete(DOC_PATH + strAssessNoCalDocName)

                    ''Threading.Thread.Sleep(2000)
                End If

            End If



            ds = pOwn.RunSQLQuery("EXEC spSelCapProcess 0, '" + dtProcessingStart.ToShortDateString + "', '" + dtProcessingEnd.ToShortDateString + String.Format("','{0}'", fac))




            If ds.Tables(0).Rows.Count > 0 Then ' if owner has no rows, facility / tanks / pipes will not have rows

                Dim flagDateSpillPreventionLastTested As Boolean

                Dim flagDateOverfillPreventionLastInspected As Boolean
                Dim flagDateSecondaryContainmentLastInspected As Boolean
                Dim flagDateElectronicDeviceInspected As Boolean
                Dim flagDateATGLastInspected As Boolean
                Dim flagDateSheerValueTest As Boolean
                Dim flagDatePipeSecondary As Boolean
                Dim flagDatePipeElectronic As Boolean
                Dim flagDateSpillPreventionLastTestedSummary As Boolean
                Dim flagDateOverfillPreventionLastInspectedSummary As Boolean
                Dim flagDateSecondaryContainmentLastInspectedSummary As Boolean
                Dim flagDateElectronicDeviceInspectedSummary As Boolean
                Dim flagDateATGLastInspectedSummary As Boolean
                Dim flagDateSheerValueTestSummary As Boolean
                Dim flagDatePipeSecondarySummary As Boolean
                Dim flagDatePipeElectronicSummary As Boolean
                Dim showTankSpillPreventionUnkown As Boolean
                Dim showTankOverfillPreventionUnkown As Boolean
                Dim showTankElectronicUnkown As Boolean
                Dim showTankSecondaryUnkown As Boolean
                Dim showTankATGUnkown As Boolean
                Dim showTankTTDateUnkown As Boolean
                Dim showTankLIInspectedUnkown As Boolean
                Dim showTankLIInstallUnkown As Boolean
                Dim showTankCPDateUnkown As Boolean
                Dim showPipeElectronicUnkown As Boolean
                Dim showPipeSecondaryUnkown As Boolean
                Dim showPipeSheerUnkown As Boolean
                Dim showPipeTTDateUnkown As Boolean
                Dim showPipeCPDateUnkown As Boolean
                Dim showPipeTermCPTestUnkown As Boolean
                Dim showPipeADDLTestDateUnkown As Boolean

                Dim enableTankSpillPrevention As Boolean
                Dim enableTankOverfillPrevention As Boolean
                Dim enableTankSecondary As Boolean
                Dim enableTankElectronic As Boolean
                Dim enableTankATG As Boolean
                Dim enablePipeSheer As Boolean
                Dim enablePipeSecondary As Boolean
                Dim enablePipeElectronic As Boolean
                Dim enableTankTTDate As Boolean
                Dim enableTankCPTestDate As Boolean
                Dim enableTankLIInspectedDate As Boolean
                Dim enablePipeALLDTestDate As Boolean
                Dim enablePipeTTDate As Boolean
                Dim enablePipeTermCPTestDate As Boolean
                Dim enablePipeCPTestDate As Boolean
                Dim flagHasCIU As Boolean
                Dim dtFinalInspected As Date
                Dim dtFinalInstall As Date

                Dim userInfoLocal As MUSTER.Info.UserInfo
                userInfoLocal = MusterContainer.AppUser.RetrieveCAEHead()
                strUser = userInfoLocal.Name
                strUserPhone = CType(userInfoLocal.PhoneNumber, String)





                ' datatables to maintain the records of the information to populate calendar
                dtCalInfo = New DataTable

                dtCalInfo.Columns.Add("FACILITY_ID", GetType(Integer))
                dtCalInfo.Columns.Add("MONTH", GetType(Integer))
                dtCalInfo.Columns.Add("INFO", GetType(String))
                dtCalInfo.Columns.Add("NAME", GetType(String))
                dtCalInfo.Columns.Add("CITY", GetType(String))

                ' datatable to save summary
                dtSummary = New DataTable

                dtSummary.Columns.Add("PROCESSING_YEAR", GetType(Integer))
                dtSummary.Columns.Add("LINE_POSITION", GetType(Integer))
                dtSummary.Columns.Add("OWN_ID", GetType(Integer))
                dtSummary.Columns.Add("OWN_NAME", GetType(String))
                dtSummary.Columns.Add("FAC_ID", GetType(Integer))
                dtSummary.Columns.Add("FAC_NAME", GetType(String))
                dtSummary.Columns.Add("FAC_ADDRESS_LINE_ONE", GetType(String))
                dtSummary.Columns.Add("FAC_CITY", GetType(String))
                dtSummary.Columns.Add("FAC_STATE", GetType(String))
                dtSummary.Columns.Add("FAC_ZIP", GetType(String))
                dtSummary.Columns.Add("DESCRIPTION", GetType(String))
                dtSummary.Columns.Add("IS_DESC_PERIODIC_TEST_REQ", GetType(Boolean))
                dtSummary.Columns.Add("IS_DESC_HEADING", GetType(Boolean))
                dtSummary.Columns.Add("IS_DESC_SUB_HEADING", GetType(Boolean))
                dtSummary.Columns.Add("CREATED_BY", GetType(String))


                'With WordApp

                '.Visible = True

                '''For Each drdsOwner In ds.Tables(0).Rows ' owner
                Dim cnt As Integer = ds.Tables(0).Rows.Count
                For i As Integer = 0 To cnt - 1 ' owner 

                    If String.Format("{0}   Preparing Yearly CAP: {1}% ", oldText, Int((((i + 1) / cnt) * 100))) <> _container.Text Then

                        _container.Text = String.Format("{0}   Preparing Yearly CAP: {1}% ", oldText, Int((((i + 1) / cnt) * 100)))

                    End If


                    drdsOwner = ds.Tables(0).Rows(i)

                    If OwnerName = String.Empty OrElse IsNumeric(OwnerName) OrElse drdsOwner("OWNERNAME").ToString.ToUpper.IndexOf(OwnerName.ToUpper) > -1 Then

                        ownID = drdsOwner("OWNER_ID").ToString


                        '  bolCanSplitSummary = True

                        ' If drdsOwner("FACCOUNT") > 4 Then
                        bolOwnerNeedsCal = True
                        ' Else
                        '    bolOwnerNeedsCal = False
                        'End If

                        'dtFac.Rows.Clear()
                        dtCalInfo.Rows.Clear()
                        slMonth = New SortedList

                        bolAddedSectionForOwner = False

                        'With docSummary

                        'docSummary.Activate()

                        'For Each drdsFac In ds.Tables(1).Select("OWNER_ID = " + drdsOwner("OWNER_ID").ToString) ' facility
                        If OwnerName <> String.Empty AndAlso fac <> String.Empty Then
                            drdsFacs = ds.Tables(1).Select(String.Format("FACILITY_ID = {0}", fac.Replace(",", " OR FACILITY_ID = ")))
                        Else
                            drdsFacs = ds.Tables(1).Select("OWNER_ID = " + drdsOwner("OWNER_ID").ToString)

                        End If
                        For j As Integer = 0 To drdsFacs.Length - 1 ' facility

                            If drdsFacs(j).Item("OWNER_ID") = drdsOwner("OWNER_ID").ToString Then

                                drdsFac = drdsFacs(j)

                                bolAddedSectionForFac = False
                                bolAddedSummaryPeriodicTesting = False

                                alGW = New ArrayList
                                alATG = New ArrayList
                                alICTT = New ArrayList
                                alSIR = New ArrayList
                                alMTG = New ArrayList
                                alPipeTT = New ArrayList
                                alEALLD1 = New ArrayList
                                alEALLD2 = New ArrayList
                                alMALLD2 = New ArrayList
                                alElectInter = New ArrayList
                                alVisInter = New ArrayList
                                alImpress = New ArrayList
                                alSO = New ArrayList
                                alTOSI = New ArrayList

                                alTnkLastTCPDate = New ArrayList
                                alTnkLinedDate = New ArrayList
                                alTnkTTDate = New ArrayList
                                alTnkICExpiresDate = New ArrayList

                                alPipeCPDate = New ArrayList
                                alPipeALLDDate = New ArrayList
                                alPipeTermCPDate = New ArrayList
                                alPipeLineUSDate = New ArrayList
                                alPipeLinePressDate = New ArrayList


                                flagDateSpillPreventionLastTested = False
                                flagDateOverfillPreventionLastInspected = False
                                flagDateSecondaryContainmentLastInspected = False
                                flagDateElectronicDeviceInspected = False
                                flagDateATGLastInspected = False
                                flagDateSheerValueTest = False
                                flagDatePipeSecondary = False
                                flagDatePipeElectronic = False
                                flagDateSpillPreventionLastTestedSummary = False
                                flagDateOverfillPreventionLastInspectedSummary = False
                                flagDateSecondaryContainmentLastInspectedSummary = False
                                flagDateElectronicDeviceInspectedSummary = False
                                flagDateATGLastInspectedSummary = False
                                flagDateSheerValueTestSummary = False
                                flagDatePipeSecondarySummary = False
                                flagDatePipeElectronicSummary = False
                                showTankSpillPreventionUnkown = False
                                showTankOverfillPreventionUnkown = False
                                showTankElectronicUnkown = False
                                showTankSecondaryUnkown = False
                                showTankATGUnkown = False
                                showPipeElectronicUnkown = False
                                showPipeSecondaryUnkown = False
                                showPipeSheerUnkown = False
                                showTankTTDateUnkown = False
                                showTankLIInspectedUnkown = False
                                showTankLIInstallUnkown = False
                                showTankCPDateUnkown = False
                                showPipeTTDateUnkown = False
                                showPipeCPDateUnkown = False
                                showPipeTermCPTestUnkown = False
                                showPipeADDLTestDateUnkown = False
                                enableTankSpillPrevention = False
                                enableTankOverfillPrevention = False
                                enableTankSecondary = False
                                enableTankElectronic = False
                                enableTankATG = False
                                enableTankTTDate = False
                                enableTankCPTestDate = False
                                enableTankLIInspectedDate = False
                                enablePipeSheer = False
                                enablePipeSecondary = False
                                enablePipeElectronic = False
                                enablePipeALLDTestDate = False
                                enablePipeTTDate = False
                                enablePipeTermCPTestDate = False
                                enablePipeCPTestDate = False
                                flagHasCIU = False
                                dtFinalInstall = CDate("01/01/0001")
                                dtFinalInspected = CDate("01/01/0001")

                                For Each drdsTank In ds.Tables(2).Select("FACILITY_ID = " + drdsFac("FACILITY_ID").ToString) ' tank

                                    ' check tank conditions
                                    ' if tank data used, set tank's, facility's and owner's MODIFIED column value to true
                                    ' and add row to dtCalInfo

                                    If (Not drdsTank("STATUS") Is DBNull.Value) Then

                                        If drdsTank("STATUS").ToString.IndexOf("Currently In Use") > -1 Or drdsTank("STATUS").ToString.IndexOf("Temporarily Out of Service Indefinitely") > -1 Then


                                            'Added by Hua Cao 11/20/08 add the 10 new fields to this report
                                            'DatespillPreventionLastTested
                                            If drdsTank("STATUS").ToString.IndexOf("Currently In Use") > -1 Then


                                                flagHasCIU = True
                                                If (Not drdsTank("Substance").ToString = "Used Oil") And (Not drdsTank("SmallDelivery").ToString = "True") Then

                                                    enableTankSpillPrevention = True

                                                    If Not drdsTank("DateSpillPreventionLastTested") Is DBNull.Value Then


                                                        dt = drdsTank("DateSpillPreventionLastTested")
                                                        dt = dt.Date
                                                        'dt = DateAdd(DateInterval.Year, 1, dt)
                                                        dt = NewDateAdd(dt, 1, processingYear)


                                                        If Not bolAddedSectionForOwner Then

                                                            AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                            bolAddedSectionForOwner = True
                                                            bolAddedSectionForFac = True
                                                            bolEnteredTextToSummary = True
                                                        ElseIf Not bolAddedSectionForFac Then

                                                            AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                            bolAddedSectionForFac = True
                                                            bolEnteredTextToSummary = True
                                                        End If

                                                        If (Not flagDateSpillPreventionLastTestedSummary) Then
                                                            strSummaryDesc = "Testing of spill containment buckets is required once every 12 months." + vbCrLf + _
                                                                                "According to our records, your last test was accomplished on " + DateAdd(DateInterval.Year, -1, dt).ToShortDateString + "."
                                                            AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                                            bolAddedSummaryPeriodicTesting = True
                                                            flagDateSpillPreventionLastTestedSummary = True
                                                            showTankSpillPreventionUnkown = True
                                                        End If


                                                        If bolOwnerNeedsCal Then

                                                            If Date.Compare(DateAdd(DateInterval.Year, -0, dt), dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, DateAdd(DateInterval.Year, -0, dt)) >= 0 Then
                                                                If Not flagDateSpillPreventionLastTested Then
                                                                    drCalInfo = dtCalInfo.NewRow
                                                                    drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                    drCalInfo("NAME") = drdsFac("NAME")
                                                                    drCalInfo("CITY") = drdsFac("CITY")
                                                                    drCalInfo("MONTH") = dt.Month
                                                                    drCalInfo("INFO") = "Testing of spill containment buckets by " + DateAdd(DateInterval.Year, -0, dt).ToShortDateString
                                                                    dtCalInfo.Rows.Add(drCalInfo)
                                                                    flagDateSpillPreventionLastTested = True
                                                                End If

                                                                If Not slMonth.Contains(dt.Month) Then
                                                                    slMonth.Add(dt.Month, dt.Month)
                                                                End If

                                                            End If ' If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then

                                                        End If ' If bolOwnerNeedsCal Then

                                                    Else


                                                        If bolOwnerNeedsCal Then

                                                            If Not flagDateSpillPreventionLastTested Then
                                                                drCalInfo = dtCalInfo.NewRow
                                                                drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                drCalInfo("NAME") = drdsFac("NAME")
                                                                drCalInfo("CITY") = drdsFac("CITY")
                                                                drCalInfo("MONTH") = 0
                                                                drCalInfo("INFO") = "Last testing date of the spill containment buckets is UNKNOWN"
                                                                dtCalInfo.Rows.Add(drCalInfo)

                                                            End If

                                                        End If ' If bolOwnerNeedsCal Then


                                                    End If


                                                    'DateOverfillPreventionLastInspected
                                                    enableTankOverfillPrevention = True

                                                    If (Not drdsTank("DateOverfillPreventionLastInspected") Is DBNull.Value) Then
                                                        dt = drdsTank("DateOverfillPreventionLastInspected")
                                                        dt = dt.Date
                                                        'dt = DateAdd(DateInterval.Year, 1, dt)
                                                        dt = NewDateAdd(dt, 1, processingYear)

                                                        If Not bolAddedSectionForOwner Then
                                                            AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                            bolAddedSectionForOwner = True
                                                            bolAddedSectionForFac = True
                                                            bolEnteredTextToSummary = True

                                                        ElseIf Not bolAddedSectionForFac Then
                                                            AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                            bolAddedSectionForFac = True
                                                            bolEnteredTextToSummary = True
                                                        End If

                                                        If (Not flagDateOverfillPreventionLastInspectedSummary) Then
                                                            strSummaryDesc = "Inspection of overfill prevention devices is required once every 12 months." + vbCrLf + _
                                                                                "According to our records, your last inspection was accomplished on " + DateAdd(DateInterval.Year, -1, dt).ToShortDateString + "."
                                                            AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)

                                                            bolAddedSummaryPeriodicTesting = True
                                                            showTankOverfillPreventionUnkown = True
                                                        End If

                                                        If bolOwnerNeedsCal Then

                                                            If Date.Compare(DateAdd(DateInterval.Year, -0, dt), dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, DateAdd(DateInterval.Year, -0, dt)) >= 0 Then

                                                                If (Not flagDateOverfillPreventionLastInspected) Then

                                                                    drCalInfo = dtCalInfo.NewRow
                                                                    drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                    drCalInfo("NAME") = drdsFac("NAME")
                                                                    drCalInfo("CITY") = drdsFac("CITY")
                                                                    drCalInfo("MONTH") = dt.Month
                                                                    drCalInfo("INFO") = "Inspection of overfill prevention devices by " + DateAdd(DateInterval.Year, -0, dt).ToShortDateString
                                                                    dtCalInfo.Rows.Add(drCalInfo)
                                                                    flagDateOverfillPreventionLastInspected = True
                                                                End If

                                                                If Not slMonth.Contains(dt.Month) Then
                                                                    slMonth.Add(dt.Month, dt.Month)
                                                                End If

                                                            End If ' If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then

                                                        End If ' If bolOwnerNeedsCal Then

                                                    Else


                                                        If bolOwnerNeedsCal Then

                                                            If (Not flagDateOverfillPreventionLastInspected) Then

                                                                drCalInfo = dtCalInfo.NewRow
                                                                drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                drCalInfo("NAME") = drdsFac("NAME")
                                                                drCalInfo("CITY") = drdsFac("CITY")
                                                                drCalInfo("MONTH") = 0
                                                                drCalInfo("INFO") = "Last inspection date of the overfill prevention devices is UNKNOWN "
                                                                dtCalInfo.Rows.Add(drCalInfo)

                                                            End If
                                                        End If ' If bolOwnerNeedsCal Then

                                                    End If


                                                    'DateElectronicDeviceInspected
                                                    If drdsTank("TankLD").ToString = "Electronic Interstitial Monitoring" Then
                                                        enableTankElectronic = True
                                                        If (Not drdsTank("DateElectronicDeviceInspected") Is DBNull.Value) Then
                                                            dt = drdsTank("DateElectronicDeviceInspected")
                                                            dt = dt.Date
                                                            'dt = DateAdd(DateInterval.Year, 1, dt)
                                                            dt = NewDateAdd(dt, 1, processingYear)
                                                            If Not bolAddedSectionForOwner Then
                                                                AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                                bolAddedSectionForOwner = True
                                                                bolAddedSectionForFac = True
                                                                bolEnteredTextToSummary = True
                                                            ElseIf Not bolAddedSectionForFac Then
                                                                AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                                bolAddedSectionForFac = True
                                                                bolEnteredTextToSummary = True
                                                            End If
                                                            If Not flagDateElectronicDeviceInspectedSummary Then
                                                                strSummaryDesc = "Testing of tank electronic interstitial monitoring devices is required once every 12 months." + vbCrLf + _
                                                                                    "According to our records, your last test was accomplished on " + DateAdd(DateInterval.Year, -1, dt).ToShortDateString + "."
                                                                AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                                                bolAddedSummaryPeriodicTesting = True
                                                                flagDateElectronicDeviceInspectedSummary = True
                                                                showTankElectronicUnkown = True
                                                            End If

                                                            If bolOwnerNeedsCal Then
                                                                If Date.Compare(DateAdd(DateInterval.Year, -0, dt), dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, DateAdd(DateInterval.Year, -0, dt)) >= 0 Then
                                                                    If Not flagDateElectronicDeviceInspected Then
                                                                        drCalInfo = dtCalInfo.NewRow
                                                                        drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                        drCalInfo("NAME") = drdsFac("NAME")
                                                                        drCalInfo("CITY") = drdsFac("CITY")
                                                                        drCalInfo("MONTH") = dt.Month
                                                                        drCalInfo("INFO") = "Testing of tank electronic interstitial monitoring devices by " + DateAdd(DateInterval.Year, -0, dt).ToShortDateString
                                                                        dtCalInfo.Rows.Add(drCalInfo)

                                                                        flagDateElectronicDeviceInspected = True
                                                                    End If
                                                                    If Not slMonth.Contains(dt.Month) Then
                                                                        slMonth.Add(dt.Month, dt.Month)
                                                                    End If


                                                                End If ' If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                            End If ' If bolOwnerNeedsCal Then

                                                        Else

                                                            If bolOwnerNeedsCal Then
                                                                If Not flagDateElectronicDeviceInspected Then

                                                                    drCalInfo = dtCalInfo.NewRow
                                                                    drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                    drCalInfo("NAME") = drdsFac("NAME")
                                                                    drCalInfo("CITY") = drdsFac("CITY")
                                                                    drCalInfo("MONTH") = 0
                                                                    drCalInfo("INFO") = "Last testing date of the tank electronic interstitial monitoring devices is UNKNOWN"
                                                                    dtCalInfo.Rows.Add(drCalInfo)

                                                                End If
                                                            End If ' If bolOwnerNeedsCal Then
                                                        End If

                                                    End If

                                                    'DateATGLastInspected
                                                    If (drdsTank("TankLD").ToString = "Automatic Tank Gauging") Then
                                                        enableTankATG = True
                                                        If (Not drdsTank("DateATGLastInspected") Is DBNull.Value) Then

                                                            dt = drdsTank("DateATGLastInspected")
                                                            dt = dt.Date
                                                            'dt = DateAdd(DateInterval.Year, 1, dt)
                                                            dt = NewDateAdd(dt, 1, processingYear)
                                                            If Not bolAddedSectionForOwner Then
                                                                AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                                bolAddedSectionForOwner = True
                                                                bolAddedSectionForFac = True
                                                                bolEnteredTextToSummary = True
                                                            ElseIf Not bolAddedSectionForFac Then
                                                                AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                                bolAddedSectionForFac = True
                                                                bolEnteredTextToSummary = True
                                                            End If
                                                            If (Not flagDateATGLastInspectedSummary) Then
                                                                strSummaryDesc = "Inspection of automatic tank gauging equipment is required once every 12 months." + vbCrLf + _
                                                                                    "According to our records, your last inspection was accomplished on " + DateAdd(DateInterval.Year, -1, dt).ToShortDateString + "."
                                                                AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                                                bolAddedSummaryPeriodicTesting = True
                                                                flagDateATGLastInspectedSummary = True
                                                                showTankATGUnkown = True
                                                            End If
                                                            If bolOwnerNeedsCal Then
                                                                If Date.Compare(DateAdd(DateInterval.Year, -0, dt), dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, DateAdd(DateInterval.Year, -0, dt)) >= 0 Then
                                                                    If (Not flagDateATGLastInspected) Then
                                                                        drCalInfo = dtCalInfo.NewRow
                                                                        drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                        drCalInfo("NAME") = drdsFac("NAME")
                                                                        drCalInfo("CITY") = drdsFac("CITY")
                                                                        drCalInfo("MONTH") = dt.Month
                                                                        drCalInfo("INFO") = "Inspection of automatic tank gauging equipment by " + DateAdd(DateInterval.Year, -0, dt).ToShortDateString
                                                                        dtCalInfo.Rows.Add(drCalInfo)

                                                                        flagDateATGLastInspected = True
                                                                    End If

                                                                    If Not slMonth.Contains(dt.Month) Then
                                                                        slMonth.Add(dt.Month, dt.Month)
                                                                    End If
                                                                End If ' If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                            End If ' If bolOwnerNeedsCal Then
                                                        Else
                                                            If bolOwnerNeedsCal Then
                                                                If (Not flagDateATGLastInspected) Then
                                                                    drCalInfo = dtCalInfo.NewRow
                                                                    drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                    drCalInfo("NAME") = drdsFac("NAME")
                                                                    drCalInfo("CITY") = drdsFac("CITY")
                                                                    drCalInfo("MONTH") = 0
                                                                    drCalInfo("INFO") = "Last inspection date of the automatic tank gauging equipment is UNKNOWN"

                                                                    dtCalInfo.Rows.Add(drCalInfo)

                                                                End If

                                                            End If
                                                        End If ' If bolOwnerNeedsCal Then
                                                    End If


                                                    ''DateSecondaryContainmentLastInspected
                                                    'If drdsTank("TankLD").ToString = "Visual Interstitial Monitoring" Then
                                                    'enableTankSecondary = True
                                                    'If (Not drdsTank("DateSecondaryContainmentLastInspected") Is DBNull.Value) Then
                                                    'dt = drdsTank("DateSecondaryContainmentLastInspected")
                                                    'dt = dt.Date
                                                    'dt = DateAdd(DateInterval.Year, 1, dt)
                                                    'dt = NewDateAdd(dt, 1, processingYear)

                                                    'If Not bolAddedSectionForOwner Then
                                                    'AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                    'bolAddedSectionForOwner = True
                                                    'bolAddedSectionForFac = True
                                                    'bolEnteredTextToSummary = True
                                                    'ElseIf Not bolAddedSectionForFac Then
                                                    'AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                    'bolAddedSectionForFac = True
                                                    'bolEnteredTextToSummary = True
                                                    'End If
                                                    'If (Not flagDateSecondaryContainmentLastInspectedSummary) Then
                                                    ' strSummaryDesc = "Inspection of the tank secondary containment sump is required once every 12 months." + vbCrLf + _
                                                    '                    "According to our records, your last inspection was accomplished on " + DateAdd(DateInterval.Year, -1, dt).ToShortDateString + "."
                                                    'AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                                    'bolAddedSummaryPeriodicTesting = True
                                                    'flagDateSecondaryContainmentLastInspectedSummary = True
                                                    'showTankSecondaryUnkown = True
                                                    'End If
                                                    'If bolOwnerNeedsCal Then
                                                    '   If Date.Compare(DateAdd(DateInterval.Year, -0, dt), dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, DateAdd(DateInterval.Year, -0, dt)) >= 0 Then
                                                    '  If (Not flagDateSecondaryContainmentLastInspected) Then
                                                    ' drCalInfo = dtCalInfo.NewRow
                                                    'drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                    'd'rCalInfo("NAME") = drdsFac("NAME")
                                                    'dr() 'CalInfo("CITY") = drdsFac("CITY")
                                                    'drCalInfo("MONTH") = dt.Month
                                                    'drCalInfo("INFO") = "Inspection of the tank secondary containment by " + DateAdd(DateInterval.Year, -0, dt).ToShortDateString
                                                    'dtCalInfo.Rows.Add(drCalInfo)
                                                    'flagDateSecondaryContainmentLastInspected = True
                                                    'End If
                                                    'If Not slMonth.Contains(dt.Month) Then
                                                    ' slMonth.Add(dt.Month, dt.Month)
                                                    'End If
                                                    ' End If ' If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                    ' End If ' If bolOwnerNeedsCal Then
                                                    'Else
                                                    '               If bolOwnerNeedsCal Then
                                                    '              If (Not flagDateSecondaryContainmentLastInspected) Then
                                                    '             drCalInfo = dtCalInfo.NewRow
                                                    '            drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                    '           drCalInfo("NAME") = drdsFac("NAME")
                                                    '          drCalInfo("CITY") = drdsFac("CITY")
                                                    '         drCalInfo("MONTH") = 0
                                                    '        drCalInfo("INFO") = "Date inspection of the tank secondary is UNKNOWN"
                                                    '
                                                    '                                               dtCalInfo.Rows.Add(drCalInfo)
                                                    '                                              flagDateSecondaryContainmentLastInspected = True
                                                    '                                         End If
                                                    '                                    End If ' If bolOwnerNeedsCal Then

                                                    '                               End If
                                                    '                          End If
                                                End If
                                            End If ' If Tank Status = Currently In Use

                                            ' LAST TCP DATE

                                            If Not drdsTank("TANKMODDESC") Is DBNull.Value Then
                                                If drdsTank("TANKMODDESC").ToString.IndexOf("Cathodically Protected") > -1 Then
                                                    enableTankCPTestDate = True
                                                    If (Not drdsTank("CP DATE") Is DBNull.Value) Then
                                                        dt = drdsTank("CP DATE")
                                                        dt = dt.Date
                                                        'dt = DateAdd(DateInterval.Year, 3, dt)
                                                        dt = NewDateAdd(dt, 3, processingYear)

                                                        If Not alTnkLastTCPDate.Contains(dt.ToShortDateString) Then
                                                            'If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                            If Not bolAddedSectionForOwner Then
                                                                'AddCAPAnnualSummaryHeading(drdsOwner("OWNERNAME").ToString, drdsFac, docSummary, bolEnteredTextToSummary, bolAddedSectionForFac, bolCanSplitSummary)
                                                                AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                                bolAddedSectionForOwner = True
                                                                bolAddedSectionForFac = True
                                                                bolEnteredTextToSummary = True
                                                            ElseIf Not bolAddedSectionForFac Then
                                                                'AddCAPAnnualSummaryHeading(drdsOwner("OWNERNAME").ToString, drdsFac, docSummary, bolEnteredTextToSummary, Not bolAddedSectionForFac, bolCanSplitSummary)
                                                                AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                                bolAddedSectionForFac = True
                                                                bolEnteredTextToSummary = True
                                                            End If
                                                            alTnkLastTCPDate.Add(dt.ToShortDateString)

                                                            strSummaryDesc = "Testing of Tank Cathodic Protection is required once every three years." + vbCrLf + _
                                                                                "According to our records, your last test was accomplished on " + DateAdd(DateInterval.Year, -3, dt).ToShortDateString + "."
                                                            AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                                            bolAddedSummaryPeriodicTesting = True
                                                            showTankCPDateUnkown = True
                                                            If bolOwnerNeedsCal Then
                                                                If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                                    drCalInfo = dtCalInfo.NewRow
                                                                    drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                    drCalInfo("NAME") = drdsFac("NAME")
                                                                    drCalInfo("CITY") = drdsFac("CITY")
                                                                    drCalInfo("MONTH") = dt.Month
                                                                    drCalInfo("INFO") = "Test tank cathodic protection by " + dt.ToShortDateString
                                                                    dtCalInfo.Rows.Add(drCalInfo)

                                                                    If Not slMonth.Contains(dt.Month) Then
                                                                        slMonth.Add(dt.Month, dt.Month)
                                                                    End If
                                                                End If ' If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                            End If ' If bolOwnerNeedsCal Then
                                                        End If
                                                    Else

                                                        If bolOwnerNeedsCal Then
                                                            drCalInfo = dtCalInfo.NewRow
                                                            drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                            drCalInfo("NAME") = drdsFac("NAME")
                                                            drCalInfo("CITY") = drdsFac("CITY")
                                                            drCalInfo("MONTH") = 0
                                                            drCalInfo("INFO") = "Last testing date of the tank cathodic protection is UNKNOWN"
                                                            dtCalInfo.Rows.Add(drCalInfo)

                                                        End If ' If bolOwnerNeedsCal Then

                                                    End If
                                                End If
                                            End If

                                            ' LINED DUE
                                            If Not drdsTank("TANKMODDESC") Is DBNull.Value Then
                                                If drdsTank("TANKMODDESC").ToString = "Lined Interior" Then
                                                    enableTankLIInspectedDate = True
                                                    dt = IIf(drdsTank("LI INSTALL") Is DBNull.Value, CDate("01/01/0001"), drdsTank("LI INSTALL"))
                                                    dt1 = IIf(drdsTank("LI INSPECTED") Is DBNull.Value, CDate("01/01/0001"), drdsTank("LI INSPECTED"))
                                                    dt = dt.Date
                                                    dt1 = dt1.Date
                                                    dt = DateAdd(DateInterval.Year, 10, dt)
                                                    dt1 = DateAdd(DateInterval.Year, 5, dt1)

                                                    If Not (Date.Compare(dt, dt1) > 0 Or Date.Compare(dt1, CDate("01/01/0001")) = 0) Then
                                                        dt = dt1
                                                    End If
                                                    If Not drdsTank("TANKMODDESC").ToString.IndexOf("Cathodically Protected/Lined Interior") > -1 Then


                                                        'If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                        If Not alTnkLinedDate.Contains(dt.ToShortDateString) Then
                                                            If Not bolAddedSectionForOwner Then
                                                                'AddCAPAnnualSummaryHeading(drdsOwner("OWNERNAME").ToString, drdsFac, docSummary, bolEnteredTextToSummary, bolAddedSectionForFac, bolCanSplitSummary)
                                                                AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                                bolAddedSectionForOwner = True
                                                                bolAddedSectionForFac = True
                                                                bolEnteredTextToSummary = True
                                                            ElseIf Not bolAddedSectionForFac Then
                                                                'AddCAPAnnualSummaryHeading(drdsOwner("OWNERNAME").ToString, drdsFac, docSummary, bolEnteredTextToSummary, Not bolAddedSectionForFac, bolCanSplitSummary)
                                                                AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                                bolAddedSectionForFac = True
                                                                bolEnteredTextToSummary = True
                                                            End If
                                                            alTnkLinedDate.Add(dt.ToShortDateString)

                                                            ' strSummaryDesc = "Inspection of the tank internal lining is required within 10 years of the date the lining was installed and once every five years thereafter." + vbCrLf + _
                                                            '                   "According to our records, the tank internal lining was installed on "
                                                            Dim dt2 As Date = IIf(drdsTank("LI INSPECTED") Is DBNull.Value, CDate("01/01/0001"), drdsTank("LI INSPECTED"))
                                                            Dim dt2Installed As Date = IIf(drdsTank("LI INSTALL") Is DBNull.Value, CDate("01/01/0001"), drdsTank("LI INSTALL"))
                                                            'If Date.Compare(dt2, CDate("01/01/0001")) = 0 Then
                                                            'dt2 = IIf(drdsTank("LI INSTALL") Is DBNull.Value, CDate("01/01/0001"), drdsTank("LI INSTALL"))
                                                            'End If
                                                            'strSummaryDesc += dt2Installed.Date.ToShortDateString + " and last inspected on " + dt2.Date.ToShortDateString
                                                            'AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                                            If Not (Date.Compare(dt2, CDate("01/01/0001")) = 0) Then
                                                                dtFinalInspected = dt2
                                                                showTankLIInspectedUnkown = True
                                                            End If
                                                            If Not (Date.Compare(dt2Installed, CDate("01/01/0001")) = 0) Then
                                                                dtFinalInstall = dt2Installed
                                                                showTankLIInstallUnkown = True
                                                            End If
                                                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                            '    bolAddedSummaryPeriodicTesting = True
                                                            '''''''''''''''''''''''''''''''''''''''''''''''''''
                                                            If bolOwnerNeedsCal Then

                                                                If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                                    drCalInfo = dtCalInfo.NewRow
                                                                    drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                    drCalInfo("NAME") = drdsFac("NAME")
                                                                    drCalInfo("CITY") = drdsFac("CITY")
                                                                    drCalInfo("MONTH") = dt.Month
                                                                    drCalInfo("INFO") = "Test tank interior lining by " + dt2.ToShortDateString
                                                                    dtCalInfo.Rows.Add(drCalInfo)

                                                                    If Not slMonth.Contains(dt2.Month) Then
                                                                        slMonth.Add(dt2.Month, dt2.Month)
                                                                    End If
                                                                ElseIf dt.Year <= dtProcessingStart.Year AndAlso drdsTank("LI INSPECTED") Is DBNull.Value Then
                                                                    drCalInfo = dtCalInfo.NewRow
                                                                    drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                    drCalInfo("NAME") = drdsFac("NAME")
                                                                    drCalInfo("CITY") = drdsFac("CITY")
                                                                    drCalInfo("MONTH") = 0
                                                                    drCalInfo("INFO") = "Date of last tank internal lining inspection is UNKNOWN"
                                                                    dtCalInfo.Rows.Add(drCalInfo)
                                                                ElseIf dt.Year <= dtProcessingStart.Year AndAlso Not drdsTank("LI INSPECTED") Is DBNull.Value AndAlso DateAdd(DateInterval.Year, 5, drdsTank("LI INSPECTED")).Year < dtProcessingStart.Year Then
                                                                    drCalInfo = dtCalInfo.NewRow
                                                                    drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                    drCalInfo("NAME") = drdsFac("NAME")
                                                                    drCalInfo("CITY") = drdsFac("CITY")
                                                                    drCalInfo("MONTH") = 0
                                                                    drCalInfo("INFO") = "Date of last tank internal lining inspection is UNKNOWN"
                                                                    dtCalInfo.Rows.Add(drCalInfo)


                                                                End If ' If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then

                                                            End If ' If bolOwnerNeedsCal Then

                                                        End If

                                                    End If
                                                End If
                                            End If

                                            ' IF CIU
                                            If drdsTank("STATUS").ToString.IndexOf("Currently In Use") > -1 Then
                                                ' TANK TT DUE / IC EXPIRES
                                                ' Show Tank TT Due only if IC Expires is false
                                                Dim showTankTTDue As Boolean = True
                                                If Not drdsTank("TANKLD") Is DBNull.Value Then
                                                    If drdsTank("TANKLD").ToString.IndexOf("Inventory Control/Precision Tightness Testing") > -1 Then
                                                        ' ICExpires
                                                        enableTankTTDate = True
                                                        dt = IIf(drdsTank("INSTALLED") Is DBNull.Value, CDate("01/01/0001"), drdsTank("INSTALLED"))
                                                        dt1 = IIf(drdsTank("TCPINSTALLDATE") Is DBNull.Value, CDate("01/01/0001"), drdsTank("TCPINSTALLDATE"))
                                                        dt = dt.Date
                                                        dt1 = dt1.Date
                                                        If Date.Compare(dt, dt1) < 0 Then
                                                            dt = dt1
                                                        End If
                                                        dt1 = IIf(drdsTank("LI INSTALL") Is DBNull.Value, CDate("01/01/0001"), drdsTank("LI INSTALL"))
                                                        dt1 = dt1.Date
                                                        If Date.Compare(dt, dt1) < 0 Then
                                                            dt = dt1
                                                        End If
                                                        dt = DateAdd(DateInterval.Year, 10, dt)
                                                        dt1 = CDate("12/22/1998")
                                                        If Date.Compare(dt, dt1) < 0 Then
                                                            dt = dt1
                                                        End If

                                                        'If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                        If Not alTnkICExpiresDate.Contains(dt.ToShortDateString) Then
                                                            If Not bolAddedSectionForOwner Then
                                                                'AddCAPAnnualSummaryHeading(drdsOwner("OWNERNAME").ToString, drdsFac, docSummary, bolEnteredTextToSummary, bolAddedSectionForFac, bolCanSplitSummary)
                                                                AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                                bolAddedSectionForOwner = True
                                                                bolAddedSectionForFac = True
                                                                bolEnteredTextToSummary = True
                                                            ElseIf Not bolAddedSectionForFac Then
                                                                'AddCAPAnnualSummaryHeading(drdsOwner("OWNERNAME").ToString, drdsFac, docSummary, bolEnteredTextToSummary, Not bolAddedSectionForFac, bolCanSplitSummary)
                                                                AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                                bolAddedSectionForFac = True
                                                                bolEnteredTextToSummary = True
                                                            End If


                                                            alTnkICExpiresDate.Add(dt.ToShortDateString)

                                                            strSummaryDesc = "Note that Inventory Control/Precision Tightness Testing is only a valid method of Tank Leak " + _
                                                                                "detection for a period of 10 years following tank installation / upgrade. " + vbCrLf + _
                                                                                    "According to our records, your PTT was installed on "

                                                            If Date.Compare(dt.ToShortDateString, CDate("12/22/1998")) = 0 Then
                                                                strSummaryDesc += dt.ToShortDateString + "."
                                                            Else
                                                                strSummaryDesc += DateAdd(DateInterval.Year, -10, dt).ToShortDateString + "."
                                                            End If

                                                            showTankTTDateUnkown = True

                                                            AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                                            'oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                                                            ''oPara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                                                            'oPara.Range.Font.Name = "Arial"
                                                            'oPara.Range.Font.Size = 10
                                                            'oPara.Range.Font.Bold = 0
                                                            'oPara.Range.Text = "Note that Inventory Control/Precision Tightness Testing is only a valid method of Tank Leak " + _
                                                            '                    "detection for a period of 10 years following tank installation / upgrade. Therefore, by no later" + _
                                                            '                    "than " + dt.ToShortDateString + " you must choose another method."
                                                            'oPara.Range.InsertParagraphAfter()

                                                            'InsertLines(1, docSummary)
                                                            bolAddedSummaryPeriodicTesting = True

                                                            If bolOwnerNeedsCal Then
                                                                If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                                    drCalInfo = dtCalInfo.NewRow
                                                                    drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                    drCalInfo("NAME") = drdsFac("NAME")
                                                                    drCalInfo("CITY") = drdsFac("CITY")
                                                                    drCalInfo("MONTH") = dt.Month
                                                                    drCalInfo("INFO") = "Please be aware that Inventory Control/Precision Tightness testing is only a valid method of Tank Leak " + _
                                                                                        "detection for a period of 10 years following tank installation / upgrade. Therefore, by no later " + _
                                                                                        "than " + dt.ToShortDateString + " you must choose another method."
                                                                    dtCalInfo.Rows.Add(drCalInfo)

                                                                    If Not slMonth.Contains(dt.Month) Then
                                                                        slMonth.Add(dt.Month, dt.Month)
                                                                    End If
                                                                End If ' If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                            End If ' If bolOwnerNeedsCal Then

                                                        End If
                                                        'End If

                                                        If showTankTTDue Then
                                                            ' TANK TT DUE
                                                            dt = IIf(drdsTank("TT DATE") Is DBNull.Value, CDate("01/01/0001"), drdsTank("TT DATE"))
                                                            dt1 = IIf(drdsTank("INSTALLED") Is DBNull.Value, CDate("01/01/0001"), drdsTank("INSTALLED"))
                                                            dt = dt.Date
                                                            dt1 = dt1.Date

                                                            If Date.Compare(dt, dt1) < 0 Then
                                                                dt = dt1
                                                            End If
                                                            dt1 = IIf(drdsTank("TCPINSTALLDATE") Is DBNull.Value, CDate("01/01/0001"), drdsTank("TCPINSTALLDATE"))
                                                            dt1 = dt1.Date
                                                            If Date.Compare(dt, dt1) < 0 Then
                                                                dt = dt1
                                                            End If
                                                            dt1 = IIf(drdsTank("LI INSTALL") Is DBNull.Value, CDate("01/01/0001"), drdsTank("LI INSTALL"))
                                                            dt1 = dt1.Date
                                                            If Date.Compare(dt, dt1) < 0 Then
                                                                dt = dt1
                                                            End If

                                                            dt = DateAdd(DateInterval.Year, 5, dt)
                                                            dt1 = CDate("12/22/1998")
                                                            If Date.Compare(dt, dt1) < 0 Then
                                                                dt = dt1
                                                            End If

                                                            'If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then

                                                            If Not alTnkTTDate.Contains(dt.ToShortDateString) Then
                                                                If Not bolAddedSectionForOwner Then
                                                                    'AddCAPAnnualSummaryHeading(drdsOwner("OWNERNAME").ToString, drdsFac, docSummary, bolEnteredTextToSummary, bolAddedSectionForFac, bolCanSplitSummary)
                                                                    AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                                    bolAddedSectionForOwner = True
                                                                    bolAddedSectionForFac = True
                                                                    bolEnteredTextToSummary = True
                                                                ElseIf Not bolAddedSectionForFac Then
                                                                    'AddCAPAnnualSummaryHeading(drdsOwner("OWNERNAME").ToString, drdsFac, docSummary, bolEnteredTextToSummary, Not bolAddedSectionForFac, bolCanSplitSummary)
                                                                    AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                                    bolAddedSectionForFac = True
                                                                    bolEnteredTextToSummary = True
                                                                End If
                                                                alTnkTTDate.Add(dt.ToShortDateString)

                                                                strSummaryDesc = "Testing of Tank Tightness is required once every five years." + vbCrLf + _
                                                                                    "According to our records, your last test was accomplished on "
                                                                If Date.Compare(dt.ToShortDateString, CDate("12/22/1998")) = 0 Then
                                                                    strSummaryDesc += dt.ToShortDateString + "."
                                                                Else
                                                                    strSummaryDesc += DateAdd(DateInterval.Year, -5, dt).ToShortDateString + "."
                                                                End If
                                                                AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                                                showTankTTDateUnkown = True
                                                                bolAddedSummaryPeriodicTesting = True

                                                                If bolOwnerNeedsCal Then
                                                                    If Date.Compare(dt.ToShortDateString, CDate("12/22/1998")) <> 0 Then
                                                                        '  dt = DateAdd(DateInterval.Year, -5, dt)
                                                                    End If
                                                                    If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                                        drCalInfo = dtCalInfo.NewRow
                                                                        drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                        drCalInfo("NAME") = drdsFac("NAME")
                                                                        drCalInfo("CITY") = drdsFac("CITY")
                                                                        drCalInfo("MONTH") = dt.Month
                                                                        drCalInfo("INFO") = "Test tank tightness by " + dt.ToShortDateString
                                                                        dtCalInfo.Rows.Add(drCalInfo)

                                                                        If Not slMonth.Contains(dt.Month) Then
                                                                            slMonth.Add(dt.Month, dt.Month)
                                                                        End If
                                                                    ElseIf dt < DateAdd(DateInterval.Year, -5, dtProcessingStart) Then
                                                                        drCalInfo = dtCalInfo.NewRow
                                                                        drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                        drCalInfo("NAME") = drdsFac("NAME")
                                                                        drCalInfo("CITY") = drdsFac("CITY")
                                                                        drCalInfo("MONTH") = 0
                                                                        drCalInfo("INFO") = "Last testing or install date of tank tightness is UNKNOWN"
                                                                        dtCalInfo.Rows.Add(drCalInfo)

                                                                    End If ' If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                                End If ' If bolOwnerNeedsCal Then
                                                            End If
                                                        End If ' if showTankTTDue

                                                        ' summary
                                                        If Not alICTT.Contains(drdsFac("FACILITY_ID")) Then
                                                            alICTT.Add(drdsFac("FACILITY_ID"))
                                                            bolAddSummaryRoutineTesting = True
                                                        End If

                                                    ElseIf drdsTank("TANKLD").ToString.IndexOf("Groundwater/Vapor Monitoring") > -1 Then
                                                        If Not alGW.Contains(drdsFac("FACILITY_ID")) Then
                                                            alGW.Add(drdsFac("FACILITY_ID"))
                                                            bolAddSummaryRoutineTesting = True
                                                        End If
                                                    ElseIf drdsTank("TANKLD").ToString.IndexOf("Automatic Tank Gauging") > -1 Then
                                                        If Not alATG.Contains(drdsFac("FACILITY_ID")) Then
                                                            alATG.Add(drdsFac("FACILITY_ID"))
                                                            bolAddSummaryRoutineTesting = True
                                                        End If
                                                    ElseIf drdsTank("TANKLD").ToString.IndexOf("Statistical Inventory Reconciliation(SIR)") > -1 Then
                                                        If Not alSIR.Contains(drdsFac("FACILITY_ID")) Then
                                                            alSIR.Add(drdsFac("FACILITY_ID"))
                                                            bolAddSummaryRoutineTesting = True
                                                        End If
                                                    ElseIf drdsTank("TANKLD").ToString.IndexOf("Manual Tank Gauging") > -1 Then
                                                        If Not alMTG.Contains(drdsFac("FACILITY_ID")) Then
                                                            alMTG.Add(drdsFac("FACILITY_ID"))
                                                            bolAddSummaryRoutineTesting = True
                                                        End If
                                                    ElseIf drdsTank("TANKLD").ToString.IndexOf("Electronic Interstitial Monitoring") > -1 Then
                                                        If Not alElectInter.Contains(drdsFac("FACILITY_ID")) Then
                                                            alElectInter.Add(drdsFac("FACILITY_ID"))
                                                            bolAddSummaryRoutineTesting = True
                                                        End If
                                                    ElseIf drdsTank("TANKLD").ToString.IndexOf("Visual Interstitial Monitoring") > -1 Then
                                                        If Not alVisInter.Contains(drdsFac("FACILITY_ID")) Then
                                                            alVisInter.Add(drdsFac("FACILITY_ID"))
                                                            bolAddSummaryRoutineTesting = True
                                                        End If
                                                    End If ' if tankld = inventory control/precision tightness testing
                                                End If ' if tank Ld is null

                                                ' need to test SO for ciu
                                                If drdsTank("SMALLDELIVERY") = False Or drdsTank("TANKEMERGEN") = False Then
                                                    If Not alSO.Contains(drdsFac("FACILITY_ID")) Then
                                                        alSO.Add(drdsFac("FACILITY_ID"))
                                                        bolAddSummaryRoutineTesting = True
                                                    End If
                                                End If

                                            Else ' tosi
                                                If Not alTOSI.Contains(drdsFac("FACILITY_ID")) Then
                                                    alTOSI.Add(drdsFac("FACILITY_ID"))
                                                    bolAddSummaryRoutineTesting = True
                                                End If
                                            End If ' if ciu

                                            ' need to test Impress for ciu / tosi
                                            If Not drdsTank("TANK CP TYPE") Is DBNull.Value Then
                                                If drdsTank("TANK CP TYPE").ToString.IndexOf("Impressed Current") > -1 Then
                                                    If Not alImpress.Contains(drdsFac("FACILITY_ID")) Then
                                                        alImpress.Add(drdsFac("FACILITY_ID"))
                                                        bolAddSummaryRoutineTesting = True
                                                    End If
                                                End If
                                            End If

                                        End If ' status is ciu / tosi
                                    End If ' status is null

                                    For Each drdsPipe In ds.Tables(3).Select("FACILITY_ID = " + drdsFac("FACILITY_ID").ToString + " AND [TANK ID] = " + drdsTank("TANK ID").ToString) ' pipe
                                        ' check pipe conditions
                                        ' if pipe modified, set MODIFIED column value to true
                                        If Not drdsPipe("STATUS") Is DBNull.Value Then
                                            If drdsPipe("STATUS").ToString.IndexOf("Currently In Use") > -1 Or drdsPipe("STATUS").ToString.IndexOf("Temporarily Out of Service Indefinitely") > -1 Then
                                                If drdsPipe("STATUS").ToString.IndexOf("Currently In Use") > -1 Then
                                                    'DateSheerValueTest
                                                    If drdsPipe("PIPE_TYPE_DESC").ToString = "Pressurized" AndAlso (Not drdsTank("TANKEMERGEN").ToString = "True") Then
                                                        enablePipeSheer = True
                                                        If (Not drdsPipe("DateSheerValueTest") Is DBNull.Value) Then
                                                            dt = drdsPipe("DateSheerValueTest")
                                                            dt = dt.Date
                                                            ' dt = DateAdd(DateInterval.Year, 1, dt)
                                                            dt = NewDateAdd(dt, 1, processingYear)

                                                            If Not bolAddedSectionForOwner Then
                                                                AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                                bolAddedSectionForOwner = True
                                                                bolAddedSectionForFac = True
                                                                bolEnteredTextToSummary = True
                                                            ElseIf Not bolAddedSectionForFac Then
                                                                AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                                bolAddedSectionForFac = True
                                                                bolEnteredTextToSummary = True
                                                            End If
                                                            If (Not flagDateSheerValueTestSummary) Then

                                                                strSummaryDesc = "Testing of pressurized piping shear valves is required once every 12 months." + vbCrLf + _
                                                                                    "According to our records, your last test was accomplished on " + DateAdd(DateInterval.Year, -1, dt).ToShortDateString + "."
                                                                AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                                                bolAddedSummaryPeriodicTesting = True
                                                                flagDateSheerValueTestSummary = True
                                                                showPipeSheerUnkown = True
                                                            End If
                                                            If bolOwnerNeedsCal Then
                                                                If Date.Compare(DateAdd(DateInterval.Year, -0, dt), dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, DateAdd(DateInterval.Year, -0, dt)) >= 0 Then
                                                                    If (Not flagDateSheerValueTest) Then
                                                                        drCalInfo = dtCalInfo.NewRow
                                                                        drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                        drCalInfo("NAME") = drdsFac("NAME")
                                                                        drCalInfo("CITY") = drdsFac("CITY")
                                                                        drCalInfo("MONTH") = dt.Month
                                                                        drCalInfo("INFO") = "Testing of pressurized piping shear valves by " + dt.ToShortDateString
                                                                        dtCalInfo.Rows.Add(drCalInfo)
                                                                        flagDateSheerValueTest = True
                                                                    End If
                                                                    If Not slMonth.Contains(dt.Month) Then
                                                                        slMonth.Add(dt.Month, dt.Month)
                                                                    End If
                                                                End If ' If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                            End If  ' If bolOwnerNeedsCal Then
                                                        Else
                                                            If bolOwnerNeedsCal Then
                                                                If (Not flagDateSheerValueTest) Then
                                                                    drCalInfo = dtCalInfo.NewRow
                                                                    drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                    drCalInfo("NAME") = drdsFac("NAME")
                                                                    drCalInfo("CITY") = drdsFac("CITY")
                                                                    drCalInfo("MONTH") = 0
                                                                    drCalInfo("INFO") = "Last testing date of the pressurized piping shear valves is UNKNOWN"
                                                                    dtCalInfo.Rows.Add(drCalInfo)
                                                                End If
                                                            End If  ' If bolOwnerNeedsCal Then

                                                        End If
                                                    End If
                                                    'pipe DateSecondaryContainmentInspect
                                                    If drdsPipe("PIPE_LD").ToString = "Visual Interstitial Monitoring" Then
                                                        enablePipeSecondary = True
                                                        If (Not drdsPipe("DateSecondaryContainmentInspect") Is DBNull.Value) And (drdsPipe("Pipe_Type_Desc").ToString.IndexOf("Pressurized") > -1 Or drdsPipe("Pipe_Type_Desc").ToString.IndexOf("U.S. Suction") > -1) And (drdsPipe("Pipe_Mod_Desc").ToString.IndexOf("Double-Walled") > -1 Or drdsPipe("Pipe_Mod_Desc").ToString.IndexOf("Double-Walled/Cathodically Protected") > -1) Then
                                                            dt = drdsPipe("DateSecondaryContainmentInspect")
                                                            dt = dt.Date
                                                            'dt = DateAdd(DateInterval.Year, 1, dt)
                                                            dt = NewDateAdd(dt, 1, processingYear)

                                                            If Not bolAddedSectionForOwner Then
                                                                AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                                bolAddedSectionForOwner = True
                                                                bolAddedSectionForFac = True
                                                                bolEnteredTextToSummary = True
                                                            ElseIf Not bolAddedSectionForFac Then
                                                                AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                                bolAddedSectionForFac = True
                                                                bolEnteredTextToSummary = True
                                                            End If
                                                            If (Not flagDatePipeSecondarySummary) Then
                                                                strSummaryDesc = "Inspection of the pipe secondary containment sumps is required once every 12 months." + vbCrLf + _
                                                                                    "According to our records, your last inspection was accomplished on " + DateAdd(DateInterval.Year, -1, dt).ToShortDateString + "."
                                                                AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                                                bolAddedSummaryPeriodicTesting = True
                                                                flagDatePipeSecondarySummary = True
                                                                showPipeSecondaryUnkown = True
                                                            End If
                                                            If bolOwnerNeedsCal Then
                                                                If Date.Compare(DateAdd(DateInterval.Year, -0, dt), dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, DateAdd(DateInterval.Year, -0, dt)) >= 0 Then
                                                                    If (Not flagDatePipeSecondary) Then
                                                                        drCalInfo = dtCalInfo.NewRow
                                                                        drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                        drCalInfo("NAME") = drdsFac("NAME")
                                                                        drCalInfo("CITY") = drdsFac("CITY")
                                                                        drCalInfo("MONTH") = dt.Month
                                                                        drCalInfo("INFO") = "Inspection of the pipe secondary containment by " + dt.ToShortDateString
                                                                        dtCalInfo.Rows.Add(drCalInfo)
                                                                        flagDatePipeSecondary = True
                                                                    End If
                                                                    If Not slMonth.Contains(dt.Month) Then
                                                                        slMonth.Add(dt.Month, dt.Month)
                                                                    End If
                                                                End If ' If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                            End If  ' If bolOwnerNeedsCal Then
                                                        ElseIf (drdsPipe("Pipe_Type_Desc").ToString.IndexOf("Pressurized") > -1 Or drdsPipe("Pipe_Type_Desc").ToString.IndexOf("U.S. Suction") > -1) And (drdsPipe("Pipe_Mod_Desc").ToString.IndexOf("Double-Walled") > -1 Or drdsPipe("Pipe_Mod_Desc").ToString.IndexOf("Double-Walled/Cathodically Protected") > -1) Then
                                                            If bolOwnerNeedsCal Then
                                                                If (Not flagDatePipeSecondary) Then
                                                                    drCalInfo = dtCalInfo.NewRow
                                                                    drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                    drCalInfo("NAME") = drdsFac("NAME")
                                                                    drCalInfo("CITY") = drdsFac("CITY")
                                                                    drCalInfo("MONTH") = 0
                                                                    drCalInfo("INFO") = "Last inspection date of the pipe secondary containment is UNKNOWN"
                                                                    dtCalInfo.Rows.Add(drCalInfo)

                                                                End If
                                                            End If  ' If bolOwnerNeedsCal Then

                                                        End If
                                                    End If
                                                    'DateElectronicDeviceInspect
                                                    If drdsPipe("PIPE_LD").ToString = "Continuous Interstitial Monitoring" Then
                                                        enablePipeElectronic = True
                                                        If (Not drdsPipe("DateElectronicDeviceInspect") Is DBNull.Value) And (drdsPipe("Pipe_Type_Desc").ToString.IndexOf("Pressurized") > -1 Or drdsPipe("Pipe_Type_Desc").ToString.IndexOf("U.S. Suction") > -1) And (drdsPipe("Pipe_Mod_Desc").ToString.IndexOf("Double-Walled") > -1 Or drdsPipe("Pipe_Mod_Desc").ToString.IndexOf("Double-Walled/Cathodically Protected") > -1) Then
                                                            dt = drdsPipe("DateElectronicDeviceInspect")
                                                            dt = dt.Date
                                                            'dt = DateAdd(DateInterval.Year, 1, dt)
                                                            dt = NewDateAdd(dt, 1, processingYear)
                                                            If Not bolAddedSectionForOwner Then
                                                                AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                                bolAddedSectionForOwner = True
                                                                bolAddedSectionForFac = True
                                                                bolEnteredTextToSummary = True
                                                            ElseIf Not bolAddedSectionForFac Then
                                                                AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                                bolAddedSectionForFac = True
                                                                bolEnteredTextToSummary = True
                                                            End If
                                                            If (Not flagDatePipeElectronicSummary) Then

                                                                strSummaryDesc = "Testing of the line electronic interstitial monitoring devices is required once every 12 months." + vbCrLf + _
                                                                                    "According to our records, your last test was accomplished on " + DateAdd(DateInterval.Year, -1, dt).ToShortDateString + "."
                                                                AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                                                bolAddedSummaryPeriodicTesting = True
                                                                flagDatePipeElectronicSummary = True
                                                                showPipeElectronicUnkown = True
                                                            End If
                                                            If bolOwnerNeedsCal Then
                                                                If Date.Compare(DateAdd(DateInterval.Year, -0, dt), dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, DateAdd(DateInterval.Year, -0, dt)) >= 0 Then
                                                                    If (Not flagDatePipeElectronic) Then
                                                                        drCalInfo = dtCalInfo.NewRow
                                                                        drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                        drCalInfo("NAME") = drdsFac("NAME")
                                                                        drCalInfo("CITY") = drdsFac("CITY")
                                                                        drCalInfo("MONTH") = dt.Month
                                                                        drCalInfo("INFO") = "Testing of the line electronic interstitial monitoring devices by " + dt.ToShortDateString
                                                                        dtCalInfo.Rows.Add(drCalInfo)
                                                                        flagDatePipeElectronic = True
                                                                    End If
                                                                    If Not slMonth.Contains(dt.Month) Then
                                                                        slMonth.Add(dt.Month, dt.Month)
                                                                    End If
                                                                End If ' If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                            End If  ' If bolOwnerNeedsCal Then

                                                        ElseIf (drdsPipe("Pipe_Type_Desc").ToString.IndexOf("Pressurized") > -1 Or drdsPipe("Pipe_Type_Desc").ToString.IndexOf("U.S. Suction") > -1) And (drdsPipe("Pipe_Mod_Desc").ToString.IndexOf("Double-Walled") > -1 Or drdsPipe("Pipe_Mod_Desc").ToString.IndexOf("Double-Walled/Cathodically Protected") > -1) Then

                                                            If bolOwnerNeedsCal Then
                                                                If (Not flagDatePipeElectronic) Then
                                                                    drCalInfo = dtCalInfo.NewRow
                                                                    drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                    drCalInfo("NAME") = drdsFac("NAME")
                                                                    drCalInfo("CITY") = drdsFac("CITY")
                                                                    drCalInfo("MONTH") = 0
                                                                    drCalInfo("INFO") = "Last testing date of the line electronic interstitial monitoring devices is UNKNOWN"
                                                                    dtCalInfo.Rows.Add(drCalInfo)

                                                                End If
                                                            End If  ' If bolOwnerNeedsCal Then

                                                        End If
                                                    End If
                                                End If ' If Pipe Status = Currently In Use

                                                ' PIPE CP DATE
                                                If drdsPipe("PIPE_MOD_DESC").ToString = "Cathodically Protected" Then

                                                    If Not drdsPipe("PIPE_MOD_DESC") Is DBNull.Value Then

                                                        enablePipeCPTestDate = True
                                                        If Not drdsPipe("CP DATE") Is DBNull.Value Then
                                                            dt = drdsPipe("CP DATE")
                                                            dt = dt.Date
                                                            'dt = DateAdd(DateInterval.Year, 3, dt)
                                                            dt = NewDateAdd(dt, 3, processingYear)

                                                            'If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                            If Not alPipeCPDate.Contains(dt.ToShortDateString) Then
                                                                If Not bolAddedSectionForOwner Then
                                                                    'AddCAPAnnualSummaryHeading(drdsOwner("OWNERNAME").ToString, drdsFac, docSummary, bolEnteredTextToSummary, bolAddedSectionForFac, bolCanSplitSummary)
                                                                    AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                                    bolAddedSectionForOwner = True
                                                                    bolAddedSectionForFac = True
                                                                    bolEnteredTextToSummary = True
                                                                ElseIf Not bolAddedSectionForFac Then
                                                                    'AddCAPAnnualSummaryHeading(drdsOwner("OWNERNAME").ToString, drdsFac, docSummary, bolEnteredTextToSummary, Not bolAddedSectionForFac, bolCanSplitSummary)
                                                                    AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                                    bolAddedSectionForFac = True
                                                                    bolEnteredTextToSummary = True
                                                                End If
                                                                alPipeCPDate.Add(dt.ToShortDateString)

                                                                strSummaryDesc = "Testing of Piping Cathodic Protection is required once every three years." + vbCrLf + _
                                                                                    "According to our records, your last test was accomplished on " + DateAdd(DateInterval.Year, -3, dt).ToShortDateString + "."
                                                                AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)

                                                                bolAddedSummaryPeriodicTesting = True
                                                                showPipeCPDateUnkown = True
                                                                If bolOwnerNeedsCal Then
                                                                    If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                                        drCalInfo = dtCalInfo.NewRow
                                                                        drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                        drCalInfo("NAME") = drdsFac("NAME")
                                                                        drCalInfo("CITY") = drdsFac("CITY")
                                                                        drCalInfo("MONTH") = dt.Month
                                                                        drCalInfo("INFO") = "Test piping cathodic protection by " + dt.ToShortDateString
                                                                        dtCalInfo.Rows.Add(drCalInfo)

                                                                        If Not slMonth.Contains(dt.Month) Then
                                                                            slMonth.Add(dt.Month, dt.Month)
                                                                        End If
                                                                    End If ' If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                                End If  ' If bolOwnerNeedsCal Then
                                                            End If
                                                        Else
                                                            If bolOwnerNeedsCal Then

                                                                drCalInfo = dtCalInfo.NewRow
                                                                drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                drCalInfo("NAME") = drdsFac("NAME")
                                                                drCalInfo("CITY") = drdsFac("CITY")
                                                                drCalInfo("MONTH") = 0
                                                                drCalInfo("INFO") = "Last testing date of the piping cathodic protection is UNKNOWN"
                                                                dtCalInfo.Rows.Add(drCalInfo)

                                                            End If  ' If bolOwnerNeedsCal Then


                                                        End If

                                                    End If
                                                End If

                                                ' PIPE TERM CP DATE
                                                Dim bolContinue As Boolean = False
                                                If drdsPipe("TERMINATION_TYPE_DISP").ToString = "611" Or drdsPipe("TERMINATION_TYPE_TANK").ToString = "610" Then
                                                    enablePipeTermCPTestDate = True
                                                End If
                                                If Not drdsPipe("DISP CP TYPE") Is DBNull.Value Then
                                                    If drdsPipe("DISP CP TYPE").ToString.IndexOf("Cathodically Protected") > -1 Then
                                                        If Not drdsPipe("TERM CP TEST") Is DBNull.Value Then
                                                            bolContinue = True
                                                        Else
                                                            drCalInfo = dtCalInfo.NewRow
                                                            drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                            drCalInfo("NAME") = drdsFac("NAME")
                                                            drCalInfo("CITY") = drdsFac("CITY")
                                                            drCalInfo("MONTH") = 0
                                                            drCalInfo("INFO") = "Last test date of termination cathodic protection is UNKNOWN"
                                                            dtCalInfo.Rows.Add(drCalInfo)

                                                        End If
                                                    End If
                                                End If
                                                If Not bolContinue Then
                                                    If Not drdsPipe("TANK CP TYPE") Is DBNull.Value Then
                                                        If drdsPipe("TANK CP TYPE").ToString.IndexOf("Cathodically Protected") > -1 Then
                                                            If Not drdsPipe("TERM CP TEST") Is DBNull.Value Then
                                                                bolContinue = True
                                                            Else

                                                                drCalInfo = dtCalInfo.NewRow
                                                                drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                drCalInfo("NAME") = drdsFac("NAME")
                                                                drCalInfo("CITY") = drdsFac("CITY")
                                                                drCalInfo("MONTH") = 0
                                                                drCalInfo("INFO") = "Last test date of termination cathodic protection is UNKNOWN"
                                                                dtCalInfo.Rows.Add(drCalInfo)

                                                            End If
                                                        End If
                                                    End If
                                                End If

                                                If bolContinue Then
                                                    dt = drdsPipe("TERM CP TEST")
                                                    dt = dt.Date
                                                    'dt = DateAdd(DateInterval.Year, 3, dt)
                                                    dt = NewDateAdd(dt, 3, processingYear)
                                                    'If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                    If Not alPipeTermCPDate.Contains(dt.ToShortDateString) Then
                                                        If Not bolAddedSectionForOwner Then
                                                            'AddCAPAnnualSummaryHeading(drdsOwner("OWNERNAME").ToString, drdsFac, docSummary, bolEnteredTextToSummary, bolAddedSectionForFac, bolCanSplitSummary)
                                                            AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                            bolAddedSectionForOwner = True
                                                            bolAddedSectionForFac = True
                                                            bolEnteredTextToSummary = True
                                                        ElseIf Not bolAddedSectionForFac Then
                                                            'AddCAPAnnualSummaryHeading(drdsOwner("OWNERNAME").ToString, drdsFac, docSummary, bolEnteredTextToSummary, Not bolAddedSectionForFac, bolCanSplitSummary)
                                                            AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                            bolAddedSectionForFac = True
                                                            bolEnteredTextToSummary = True
                                                        End If
                                                        alPipeTermCPDate.Add(dt.ToShortDateString)

                                                        strSummaryDesc = "Testing of Pipe Termination Cathodic Protection is required once every three years." + vbCrLf + _
                                                                            "According to our records, your last test was accomplished on " + DateAdd(DateInterval.Year, -3, dt).ToShortDateString + "."
                                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                                        bolAddedSummaryPeriodicTesting = True
                                                        showPipeTermCPTestUnkown = True
                                                        If bolOwnerNeedsCal Then
                                                            If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                                drCalInfo = dtCalInfo.NewRow
                                                                drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                drCalInfo("NAME") = drdsFac("NAME")
                                                                drCalInfo("CITY") = drdsFac("CITY")
                                                                drCalInfo("MONTH") = dt.Month
                                                                drCalInfo("INFO") = "Test termination cathodic protection by " + dt.ToShortDateString
                                                                dtCalInfo.Rows.Add(drCalInfo)

                                                                If Not slMonth.Contains(dt.Month) Then
                                                                    slMonth.Add(dt.Month, dt.Month)
                                                                End If
                                                            End If ' If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                        End If ' If bolOwnerNeedsCal Then
                                                    End If
                                                    'End If
                                                End If

                                                ' IF CIU
                                                If drdsPipe("STATUS").ToString.IndexOf("Currently In Use") > -1 Then
                                                    ' ALLD TEST DATE
                                                    If Not drdsPipe("ALLD_TEST") Is DBNull.Value Then
                                                        '  If drdsPipe("ALLD_TEST").ToString = "Mechanical" Then
                                                        If drdsPipe("PIPE_TYPE_DESC").ToString = "Pressurized" And (Not (drdsPipe("PIPE_LD").ToString = "Deferred")) Then
                                                            enablePipeALLDTestDate = True
                                                            If Not drdsPipe("ALLD_TEST_DATE") Is DBNull.Value Then
                                                                dt = drdsPipe("ALLD_TEST_DATE")
                                                                dt = dt.Date
                                                                'dt = DateAdd(DateInterval.Year, 1, dt)
                                                                dt = NewDateAdd(dt, 1, processingYear)

                                                                'If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                                If Not alPipeALLDDate.Contains(dt.ToShortDateString) Then
                                                                    If Not bolAddedSectionForOwner Then
                                                                        'AddCAPAnnualSummaryHeading(drdsOwner("OWNERNAME").ToString, drdsFac, docSummary, bolEnteredTextToSummary, bolAddedSectionForFac, bolCanSplitSummary)
                                                                        AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                                        bolAddedSectionForOwner = True
                                                                        bolAddedSectionForFac = True
                                                                        bolEnteredTextToSummary = True
                                                                    ElseIf Not bolAddedSectionForFac Then
                                                                        'AddCAPAnnualSummaryHeading(drdsOwner("OWNERNAME").ToString, drdsFac, docSummary, bolEnteredTextToSummary, Not bolAddedSectionForFac, bolCanSplitSummary)
                                                                        AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                                        bolAddedSectionForFac = True
                                                                        bolEnteredTextToSummary = True
                                                                    End If
                                                                    alPipeALLDDate.Add(dt.ToShortDateString)

                                                                    strSummaryDesc = "Testing of Pipe Automatic Line Leak Detector is required once every 12 months." + vbCrLf + _
                                                                                        "According to our records, your last test was accomplished on " + DateAdd(DateInterval.Year, -1, dt).ToShortDateString + "."
                                                                    AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                                                    bolAddedSummaryPeriodicTesting = True
                                                                    showPipeADDLTestDateUnkown = True
                                                                    If bolOwnerNeedsCal Then
                                                                        If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                                            drCalInfo = dtCalInfo.NewRow
                                                                            drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                            drCalInfo("NAME") = drdsFac("NAME")
                                                                            drCalInfo("CITY") = drdsFac("CITY")
                                                                            drCalInfo("MONTH") = dt.Month
                                                                            drCalInfo("INFO") = "Test automatic line leak detector by " + dt.ToShortDateString
                                                                            dtCalInfo.Rows.Add(drCalInfo)

                                                                            If Not slMonth.Contains(dt.Month) Then
                                                                                slMonth.Add(dt.Month, dt.Month)
                                                                            End If
                                                                        End If ' If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                                    End If ' If bolOwnerNeedsCal Then
                                                                End If
                                                                'End If
                                                            Else

                                                                If bolOwnerNeedsCal Then

                                                                    drCalInfo = dtCalInfo.NewRow
                                                                    drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                    drCalInfo("NAME") = drdsFac("NAME")
                                                                    drCalInfo("CITY") = drdsFac("CITY")
                                                                    drCalInfo("MONTH") = 0
                                                                    drCalInfo("INFO") = "Last testing date of the pipe's automatic line leak detector is UNKNOWN"
                                                                    dtCalInfo.Rows.Add(drCalInfo)

                                                                End If ' If bolOwnerNeedsCal Then

                                                            End If
                                                        End If
                                                    End If

                                                    ' PIPE LINE
                                                    If Not drdsPipe("PIPE_LD") Is DBNull.Value Then
                                                        If drdsPipe("PIPE_LD").ToString = "Line Tightness Testing" Then
                                                            enablePipeTTDate = True
                                                            If Not drdsPipe("PIPE_TYPE_DESC") Is DBNull.Value Then
                                                                If Not drdsPipe("TT DATE") Is DBNull.Value Then

                                                                    ' US SUCTION
                                                                    If drdsPipe("PIPE_TYPE_DESC").ToString = "U.S. Suction" Then
                                                                        dt = drdsPipe("TT DATE")
                                                                        dt = dt.Date
                                                                        '      dt = DateAdd(DateInterval.Year, 3, dt)
                                                                        dt = NewDateAdd(dt, 3, processingYear)

                                                                        'If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                                        If Not alPipeLineUSDate.Contains(dt.ToShortDateString) Then
                                                                            If Not bolAddedSectionForOwner Then
                                                                                'AddCAPAnnualSummaryHeading(drdsOwner("OWNERNAME").ToString, drdsFac, docSummary, bolEnteredTextToSummary, bolAddedSectionForFac, bolCanSplitSummary)
                                                                                AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                                                bolAddedSectionForOwner = True
                                                                                bolAddedSectionForFac = True
                                                                                bolEnteredTextToSummary = True
                                                                            ElseIf Not bolAddedSectionForFac Then
                                                                                'AddCAPAnnualSummaryHeading(drdsOwner("OWNERNAME").ToString, drdsFac, docSummary, bolEnteredTextToSummary, Not bolAddedSectionForFac, bolCanSplitSummary)
                                                                                AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                                                bolAddedSectionForFac = True
                                                                                bolEnteredTextToSummary = True
                                                                            End If
                                                                            alPipeLineUSDate.Add(dt.ToShortDateString)

                                                                            strSummaryDesc = "Testing of line tightness is required once every three years." + vbCrLf + _
                                                                                                "According to our records, your last test was accomplished on " + DateAdd(DateInterval.Year, -3, dt).ToShortDateString
                                                                            AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                                                            bolAddedSummaryPeriodicTesting = True
                                                                            showPipeTTDateUnkown = True
                                                                            If bolOwnerNeedsCal Then
                                                                                If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                                                    drCalInfo = dtCalInfo.NewRow
                                                                                    drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                                    drCalInfo("NAME") = drdsFac("NAME")
                                                                                    drCalInfo("CITY") = drdsFac("CITY")
                                                                                    drCalInfo("MONTH") = dt.Month
                                                                                    drCalInfo("INFO") = "Test line tightness by " + dt.ToShortDateString
                                                                                    dtCalInfo.Rows.Add(drCalInfo)

                                                                                    If Not slMonth.Contains(dt.Month) Then
                                                                                        slMonth.Add(dt.Month, dt.Month)
                                                                                    End If
                                                                                End If ' If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                                            End If ' If bolOwnerNeedsCal Then
                                                                        End If
                                                                        'End If

                                                                        ' PRESSURIZED Pipe TT Date
                                                                    ElseIf drdsPipe("PIPE_TYPE_DESC").ToString = "Pressurized" Then
                                                                        dt = drdsPipe("TT DATE")
                                                                        dt = dt.Date
                                                                        'dt = DateAdd(DateInterval.Year, 1, dt)
                                                                        dt = NewDateAdd(dt, 1, processingYear)

                                                                        'If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                                        If Not alPipeLinePressDate.Contains(dt.ToShortDateString) Then
                                                                            If Not bolAddedSectionForOwner Then
                                                                                'AddCAPAnnualSummaryHeading(drdsOwner("OWNERNAME").ToString, drdsFac, docSummary, bolEnteredTextToSummary, bolAddedSectionForFac, bolCanSplitSummary)
                                                                                AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                                                bolAddedSectionForOwner = True
                                                                                bolAddedSectionForFac = True
                                                                                bolEnteredTextToSummary = True
                                                                            ElseIf Not bolAddedSectionForFac Then
                                                                                'AddCAPAnnualSummaryHeading(drdsOwner("OWNERNAME").ToString, drdsFac, docSummary, bolEnteredTextToSummary, Not bolAddedSectionForFac, bolCanSplitSummary)
                                                                                AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                                                                bolAddedSectionForFac = True
                                                                                bolEnteredTextToSummary = True
                                                                            End If
                                                                            alPipeLinePressDate.Add(dt.ToShortDateString)

                                                                            strSummaryDesc = "Testing of line tightness is required once every 12 months." + vbCrLf + _
                                                                                                "According to our records, your last test was accomplished on " + DateAdd(DateInterval.Year, -1, dt).ToShortDateString + "."
                                                                            AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                                                            bolAddedSummaryPeriodicTesting = True
                                                                            showPipeTTDateUnkown = True
                                                                            If bolOwnerNeedsCal Then
                                                                                If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                                                    drCalInfo = dtCalInfo.NewRow
                                                                                    drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                                    drCalInfo("NAME") = drdsFac("NAME")
                                                                                    drCalInfo("CITY") = drdsFac("CITY")
                                                                                    drCalInfo("MONTH") = dt.Month
                                                                                    drCalInfo("INFO") = "Test line tightness by " + dt.ToShortDateString
                                                                                    dtCalInfo.Rows.Add(drCalInfo)

                                                                                    If Not slMonth.Contains(dt.Month) Then
                                                                                        slMonth.Add(dt.Month, dt.Month)
                                                                                    End If
                                                                                End If
                                                                            End If ' If bolOwnerNeedsCal Then

                                                                        End If ' If Not alPipeLinePressDate.Contains(dt.ToShortDateString) Then
                                                                        'End If

                                                                    End If ' If drdsPipe("PIPE_TYPE_DESC").ToString = "U.S. Suction" Then
                                                                Else

                                                                    ' US SUCTION
                                                                    If drdsPipe("PIPE_TYPE_DESC").ToString = "U.S. Suction" Then

                                                                        If bolOwnerNeedsCal Then

                                                                            drCalInfo = dtCalInfo.NewRow
                                                                            drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                            drCalInfo("NAME") = drdsFac("NAME")
                                                                            drCalInfo("CITY") = drdsFac("CITY")
                                                                            drCalInfo("MONTH") = 0
                                                                            drCalInfo("INFO") = "Last testing date of the U.S. sunction line tightness is UNKNOWN"
                                                                            dtCalInfo.Rows.Add(drCalInfo)

                                                                        End If ' If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                                        'End If

                                                                        ' PRESSURIZED Pipe TT Date
                                                                    ElseIf drdsPipe("PIPE_TYPE_DESC").ToString = "Pressurized" Then


                                                                        If bolOwnerNeedsCal Then

                                                                            drCalInfo = dtCalInfo.NewRow
                                                                            drCalInfo("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                                            drCalInfo("NAME") = drdsFac("NAME")
                                                                            drCalInfo("CITY") = drdsFac("CITY")
                                                                            drCalInfo("MONTH") = 0
                                                                            drCalInfo("INFO") = "Last testing date of the pressurized line tightness is UNKNOWN"
                                                                            dtCalInfo.Rows.Add(drCalInfo)

                                                                        End If ' If bolOwnerNeedsCal Then

                                                                        'End If

                                                                    End If ' If drdsPipe("PIPE_TYPE_DESC").ToString = "U.S. Suction" Then

                                                                End If ' If Not drdsPipe("TT DATE") Is DBNull.Value Then

                                                            End If ' If Not drdsPipe("PIPE_TYPE_DESC") Is DBNull.Value Then

                                                            ' summary
                                                            If Not alPipeTT.Contains(drdsFac("FACILITY_ID")) Then
                                                                alPipeTT.Add(drdsFac("FACILITY_ID"))
                                                                bolAddSummaryRoutineTesting = True
                                                            End If

                                                        ElseIf drdsPipe("PIPE_LD").ToString.IndexOf("Statistical Inventory Reconciliation(SIR)") > -1 Then
                                                            If Not alSIR.Contains(drdsFac("FACILITY_ID")) Then
                                                                alSIR.Add(drdsFac("FACILITY_ID"))
                                                                bolAddSummaryRoutineTesting = True
                                                            End If
                                                        ElseIf drdsPipe("PIPE_LD").ToString.IndexOf("Electronic ALLD with 0.2 Test") > -1 Then
                                                            If Not alEALLD1.Contains(drdsFac("FACILITY_ID")) Then
                                                                alEALLD1.Add(drdsFac("FACILITY_ID"))
                                                                bolAddSummaryRoutineTesting = True
                                                            End If
                                                        ElseIf drdsPipe("PIPE_LD").ToString.IndexOf("Continuous Interstitial Monitoring") > -1 Then
                                                            If Not alElectInter.Contains(drdsFac("FACILITY_ID")) Then
                                                                alElectInter.Add(drdsFac("FACILITY_ID"))
                                                                bolAddSummaryRoutineTesting = True
                                                            End If
                                                        ElseIf drdsPipe("PIPE_LD").ToString.IndexOf("Groundwater/Vapor Monitoring") > -1 Then
                                                            If Not alGW.Contains(drdsFac("FACILITY_ID")) Then
                                                                alGW.Add(drdsFac("FACILITY_ID"))
                                                                bolAddSummaryRoutineTesting = True
                                                            End If
                                                        ElseIf drdsPipe("PIPE_LD").ToString.IndexOf("Visual Interstitial Monitoring") > -1 Then
                                                            If Not alVisInter.Contains(drdsFac("FACILITY_ID")) Then
                                                                alVisInter.Add(drdsFac("FACILITY_ID"))
                                                                bolAddSummaryRoutineTesting = True
                                                            End If
                                                        End If ' If drdsPipe("PIPE_LD").ToString = "Line Tightness Testing" Then

                                                    End If ' If Not drdsPipe("PIPE_LD") Is DBNull.Value Then

                                                    ' summary
                                                    If Not drdsPipe("ALLD_TEST") Is DBNull.Value Then
                                                        If drdsPipe("ALLD_TEST").ToString.IndexOf("Electronic") > -1 Then
                                                            If Not alEALLD2.Contains(drdsFac("FACILITY_ID")) Then
                                                                alEALLD2.Add(drdsFac("FACILITY_ID"))
                                                                bolAddSummaryRoutineTesting = True
                                                            End If
                                                        ElseIf drdsPipe("ALLD_TEST").ToString.IndexOf("Mechanical") > -1 Then
                                                            If Not alMALLD2.Contains(drdsFac("FACILITY_ID")) Then
                                                                alMALLD2.Add(drdsFac("FACILITY_ID"))
                                                                bolAddSummaryRoutineTesting = True
                                                            End If
                                                        End If
                                                    End If

                                                End If ' if ciu

                                                ' need to test Impress for ciu / tosi
                                                If Not drdsPipe("PIPE CP TYPE") Is DBNull.Value Then
                                                    If drdsPipe("PIPE CP TYPE").ToString.IndexOf("Impressed Current") > -1 Then
                                                        If Not alImpress.Contains(drdsFac("FACILITY_ID")) Then
                                                            alImpress.Add(drdsFac("FACILITY_ID"))
                                                            bolAddSummaryRoutineTesting = True
                                                        End If
                                                    End If
                                                End If

                                            End If ' if status is ciu / tosi
                                        End If ' status is null

                                    Next ' pipe

                                Next ' tank

                                'put unknown date is all NULL on cap dates
                                If (Not bolAddedSummaryPeriodicTesting) Or (Not showTankSpillPreventionUnkown) Or (Not showTankOverfillPreventionUnkown) Or (Not showTankSecondaryUnkown) Or (Not showTankElectronicUnkown) Or (Not showTankATGUnkown) Or (Not showTankTTDateUnkown) Or ((Not showTankLIInspectedUnkown) Or (Not showTankLIInstallUnkown) Or showTankLIInspectedUnkown Or showTankLIInstallUnkown) Or (Not showTankCPDateUnkown) Or (Not showPipeSheerUnkown) Or (Not showPipeSecondaryUnkown) Or (Not showPipeElectronicUnkown) Or (Not showPipeTTDateUnkown) Or (Not showPipeCPDateUnkown) Or (Not showPipeADDLTestDateUnkown) Or (Not showPipeTermCPTestUnkown) Then
                                    bolAddedSummaryPeriodicTesting = True
                                    If Not bolAddedSectionForOwner Then
                                        'AddCAPAnnualSummaryHeading(drdsOwner("OWNERNAME").ToString, drdsFac, docSummary, bolEnteredTextToSummary, bolAddedSectionForFac, bolCanSplitSummary)
                                        AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                        bolAddedSectionForOwner = True
                                        bolAddedSectionForFac = True
                                        bolEnteredTextToSummary = True
                                    ElseIf Not bolAddedSectionForFac Then
                                        'AddCAPAnnualSummaryHeading(drdsOwner("OWNERNAME").ToString, drdsFac, docSummary, bolEnteredTextToSummary, Not bolAddedSectionForFac, bolCanSplitSummary)
                                        AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                        bolAddedSectionForFac = True
                                        bolEnteredTextToSummary = True
                                    End If
                                    If (Not showTankTTDateUnkown) And flagHasCIU And enableTankTTDate Then
                                        strSummaryDesc = "Testing of Tank Tightness is required once every five years." + vbCrLf + _
                                                                                      "According to our records, your last test was accomplished on unknown date."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)

                                    End If
                                    If flagHasCIU And enableTankLIInspectedDate Then
                                        If (Not showTankLIInstallUnkown) And (Not showTankLIInspectedUnkown) Then
                                            strSummaryDesc = "Inspection of the tank internal lining is required within 10 years of the installation date and once every five years thereafter." + vbCrLf + _
                                                                                      "According to our records, the tank internal lining was installed on unknown date and last inspected on unknown date."
                                            AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                        End If
                                        If (showTankLIInstallUnkown) And (Not showTankLIInspectedUnkown) Then
                                            strSummaryDesc = "Inspection of the tank internal lining is required within 10 years of the installation date and once every five years thereafter." + vbCrLf + _
                                                                                                                              "According to our records, the tank internal lining was installed on " + dtFinalInstall.Date.ToShortDateString + " and last inspected on unknown date."
                                            AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                        End If
                                        If (Not showTankLIInstallUnkown) And (showTankLIInspectedUnkown) Then
                                            strSummaryDesc = "Inspection of the tank internal lining is required within 10 years of the installation date and once every five years thereafter." + vbCrLf + _
                                                                                                                              "According to our records, the tank internal lining was installed on unknown date and last inspected on " + dtFinalInspected.Date.ToShortDateString + "."
                                            AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                        End If
                                        If showTankLIInstallUnkown And showTankLIInspectedUnkown Then
                                            strSummaryDesc = "Inspection of the tank internal lining is required within 10 years of the installation date and once every five years thereafter." + vbCrLf + _
                                                                                                                              "According to our records, the tank internal lining was installed on " + dtFinalInstall.Date.ToShortDateString + " and last inspected on " + dtFinalInspected.Date.ToShortDateString + "."
                                            AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                        End If
                                    End If

                                    If (Not showTankSpillPreventionUnkown) And enableTankSpillPrevention And flagHasCIU Then
                                        strSummaryDesc = "Testing of spill containment buckets is required once every 12 months." + vbCrLf + _
                                                                                                     "According to our records, your last test was accomplished on unknown date."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                    End If

                                    If (Not showTankOverfillPreventionUnkown) And enableTankOverfillPrevention And flagHasCIU Then
                                        strSummaryDesc = "Inspection of overfill prevention devices is required once every 12 months." + vbCrLf + _
                                                                                                     "According to our records, your last inspection was accomplished on unknown date."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                    End If

                                    If (Not showTankSecondaryUnkown) And enableTankSecondary And flagHasCIU Then
                                        strSummaryDesc = "Inspection of the tank secondary containment is required once every 12 months." + vbCrLf + _
                                                                                                     "According to our records, your last inspection was accomplished on unknown date."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                    End If

                                    If (Not showTankElectronicUnkown) And enableTankElectronic And flagHasCIU Then
                                        strSummaryDesc = "Testing of tank electronic interstitial monitoring devices is required once every 12 months." + vbCrLf + _
                                                                                                     "According to our records, your last test was accomplished on unknown date."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                    End If

                                    If (Not showTankATGUnkown) And enableTankATG And flagHasCIU Then
                                        strSummaryDesc = "Inspection of automatic tank gauging equipment is required once every 12 months." + vbCrLf + _
                                                                                                     "According to our records, your last inspection was accomplished on unknown date."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                    End If

                                    If (Not showTankCPDateUnkown) And enableTankCPTestDate And (flagHasCIU Or alTOSI.Contains(drdsFac("FACILITY_ID"))) Then
                                        strSummaryDesc = "Testing of Tank Cathodic Protection is required once every three years." + vbCrLf + _
                                                                                                     "According to our records, your last test was accomplished on unknown date."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                    End If

                                    If (Not showPipeSheerUnkown) And flagHasCIU And enablePipeSheer Then
                                        strSummaryDesc = "Testing of pressurized piping shear valves is required once every 12 months." + vbCrLf + _
                                                                            "According to our records, your last test was accomplished on unknown date."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                    End If

                                    If (Not showPipeSecondaryUnkown) And flagHasCIU And enablePipeSecondary Then
                                        strSummaryDesc = "Inspection of the pipe secondary containment is required once every 12 months." + vbCrLf + _
                                                                            "According to our records, your last inspection was accomplished on unknown date."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                    End If

                                    If (Not showPipeElectronicUnkown) And flagHasCIU And enablePipeElectronic Then
                                        strSummaryDesc = "Testing of the line electronic interstitial monitoring devices is required once every 12 months." + vbCrLf + _
                                                                            "According to our records, your last test was accomplished on unknown date."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                    End If

                                    If (Not showPipeTTDateUnkown) And flagHasCIU And enablePipeTTDate Then
                                        strSummaryDesc = "Testing of Line Tightness is required once every three years." + vbCrLf + _
                                                                                      "According to our records, your last test was accomplished on unknown date."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                    End If
                                    If (Not showPipeCPDateUnkown) And enablePipeCPTestDate And (flagHasCIU Or alTOSI.Contains(drdsFac("FACILITY_ID"))) Then
                                        strSummaryDesc = "Testing of Piping Cathodic Protection is required once every three years." + vbCrLf + _
                                                                                       "According to our records, your last test was accomplished on unknown date."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                    End If
                                    If (Not showPipeTermCPTestUnkown) And flagHasCIU And enablePipeTermCPTestDate Then
                                        strSummaryDesc = "Testing of Pipe Termination Cathodic Protection is required once every three years." + vbCrLf + _
                                                                            "According to our records, your last test was accomplished on unknown date."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                    End If
                                    If (Not showPipeADDLTestDateUnkown) And flagHasCIU And enablePipeALLDTestDate Then
                                        strSummaryDesc = "Testing of Pipe Automatic Line Leak Detector is required once every 12 months." + vbCrLf + _
                                                                                           "According to our records, your last test was accomplished on unknown date."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                    End If
                                End If
                                ' to add none in periodic test req
                                If Not bolAddedSummaryPeriodicTesting Then
                                    bolAddedSummaryPeriodicTesting = True
                                    If Not bolAddedSectionForOwner Then
                                        'AddCAPAnnualSummaryHeading(drdsOwner("OWNERNAME").ToString, drdsFac, docSummary, bolEnteredTextToSummary, bolAddedSectionForFac, bolCanSplitSummary)
                                        AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                        bolAddedSectionForOwner = True
                                        bolAddedSectionForFac = True
                                        bolEnteredTextToSummary = True
                                    ElseIf Not bolAddedSectionForFac Then
                                        'AddCAPAnnualSummaryHeading(drdsOwner("OWNERNAME").ToString, drdsFac, docSummary, bolEnteredTextToSummary, Not bolAddedSectionForFac, bolCanSplitSummary)
                                        AddCAPAnnualSummaryHeading(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, True)
                                        bolAddedSectionForFac = True
                                        bolEnteredTextToSummary = True
                                    End If
                                    strSummaryDesc = "There is no periodic testing required by the UST regulations for this facility. However, you should ensure that the manufacturer of your equipment does not require that any periodic tests be conducted."
                                    AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, True, False, False)
                                End If

                                ' add Routine Testing and Monitoring Requirements
                                If bolAddedSummaryPeriodicTesting Then
                                    strSummaryDesc = "Routine Operation and Maintenance Requirements"
                                    AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, True, False)
                                    'oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                                    ''oPara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                                    'oPara.Range.Font.Name = "Arial"
                                    'oPara.Range.Font.Size = 10
                                    'oPara.Range.Font.Bold = 1
                                    'oPara.Range.Text = "Routine Testing and Monitoring Requirements"
                                    'oPara.Range.InsertParagraphAfter()

                                    'InsertLines(1, docSummary)

                                    Dim bolAddNone As Boolean = True

                                    If alGW.Contains(drdsFac("FACILITY_ID")) Then
                                        strSummaryDesc = "Groundwater/Vapor Monitoring as a method of tank and/or piping leak detection"
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, True)

                                        strSummaryDesc = "You must check your monitoring wells once every 30 days and keep a written record of previous 12 months."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        ' strSummaryDesc = "Maintain the previous 12 months of records"
                                        'AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "If the wells contain water, visually inspect the water for product contained in tanks."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "If the wells are dry, measure vapors with instrument capable of detecting product stored in tank."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Report a possible leak to UST Branch any time there is 1/8 inch or more of free phase product or 'high' vapors."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        bolAddNone = False
                                    End If

                                    If alATG.Contains(drdsFac("FACILITY_ID")) Then
                                        strSummaryDesc = "Automatic Tank Gauging as a method of tank leak detection"
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, True)
                                        strSummaryDesc = "The tank gauge must be programmed to perform a leak test at least once every 30 days."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "The leak test may be at 0.1 or 0.2 gallons per hour."

                                        strSummaryDesc = "Maintain the previous 12 months of leak test records either by printout or electronically."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Report a possible leak to UST Branch any time you are not able to obtain a passing leak test."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        bolAddNone = False
                                    End If

                                    If alICTT.Contains(drdsFac("FACILITY_ID")) Then
                                        strSummaryDesc = "Inventory Control/Tightness Testing as a method of tank leak detection"
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, True)

                                        ' strSummaryDesc = "Maintain records for every day the tanks are in service."
                                        'AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        '  strSummaryDesc = "Measurements must be made to nearest 1/8 inch."
                                        ' AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Overage/shortage must be calculated for every day tanks are in service."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Cumulative overage/shortage must be calculated for the month and compared with the allowable variance."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Written records must show that monthly variance was compared with allowable variance."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        'strSummaryDesc = "Written records must show that tanks were checked for water at least once during the month."
                                        'AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        ' strSummaryDesc = "Maintain copy of last tank precision tightness test until next test is conducted."
                                        'AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)


                                        strSummaryDesc = "Maintain the previous 12 months of Inventory Control records."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Report to UST Branch any time the records are out of tolerance for two consecutive months."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        bolAddNone = False
                                    End If

                                    If alSIR.Contains(drdsFac("FACILITY_ID")) Then
                                        strSummaryDesc = "Statistical Inventory Control as a method of tank and/or piping leak detection"
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, True)

                                        '     strSummaryDesc = "Records must be maintained in UST Branch format."
                                        '     AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Maintain the previous 12 months of records."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        '                                    strSummaryDesc = "Submit annual summary in required format to UST Branch by February of each year."
                                        '                                    AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Two consecutive 'inconclusives' or any 'fail' requires that a precision tightness test be conducted."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Maintain the results of any investigations conducted in response to a monthly result of 'inconclusive' or 'fail'."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Report to UST Branch a possible leak whenever an 'inconclusive' or 'fail' is declared for any month."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        bolAddNone = False
                                    End If

                                    If alMTG.Contains(drdsFac("FACILITY_ID")) Then
                                        strSummaryDesc = "Manual Tank Gauging as a method of tank leak detection"
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, True)

                                        strSummaryDesc = "Tank must be taken out of service each week for a period of at least 36 hours."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Tank contents must be measured to the nearest 1/8 inch."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Written records must indicate that tank was checked for water at least once during each month."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Records must show that weekly variance has been compared with allowed standard."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Weekly variances must have been averaged and compared with the allowed monthly standard."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Report a possible leak whenever monthly variance is not within the allowed standard."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        bolAddNone = False
                                    End If

                                    If alPipeTT.Contains(drdsFac("FACILITY_ID")) Then

                                        strSummaryDesc = "Periodic Precision Tightness Testing as a method of piping leak detection"
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, True)


                                        strSummaryDesc = "Pressurized piping must have precision test conducted once every 12 months."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Suction piping that has a check valve located at the tank must have a test once every 3 years."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Report to UST Branch a possible leak any time a line fails a precision tightness test."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        bolAddNone = False
                                    End If

                                    If alEALLD1.Contains(drdsFac("FACILITY_ID")) Then

                                        strSummaryDesc = "Electronic Automatic Line Leak Detectors as a primary method of piping leak detection"
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, True)

                                        'strSummaryDesc = "Ensure that devices are installed and operating in accordance with the manufacturers requirements."
                                        'AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Device may be programmed to perform either an annual 0.1 gph test or a monthly 0.2 gph test."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Maintain previous 12 months of test records if conducting 0.2 gallon per hour test."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        '  strSummaryDesc = "Maintain record of annual 0.1 gallon test until next 0.1 gallon test is accomplished."
                                        ' AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Report to UST Branch any time the leak detector indicates a possible leak."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        bolAddNone = False
                                    End If

                                    If alEALLD2.Contains(drdsFac("FACILITY_ID")) Then
                                        strSummaryDesc = "Electronic Automatic Line Leak Detectors as a secondary method of piping leak detection"
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, True)

                                        strSummaryDesc = "Electronic line leak detectors must be tested for proper functionality every 12 months."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Device must be capable of detecting a leak equivalent to 3 gallons per hour at 10 psi."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)
                                        ' strSummaryDesc = "Maintain records of testing until next test is done."
                                        'AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Report to UST Branch any time a leak detector 'trips' and cannot be reset."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        bolAddNone = False
                                    End If

                                    If alMALLD2.Contains(drdsFac("FACILITY_ID")) Then

                                        strSummaryDesc = "Mechanical Automatic Line Leak Detectors as a secondary method of piping leak detection"
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, True)

                                        strSummaryDesc = "Mechanical line leak detectors must be tested for proper functionality every 12 months."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Device must be capable of detecting a leak equivalent to 3 gallons per hour at 10 psi."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        'strSummaryDesc = "Maintain records of testing until next text is done."
                                        'AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Report to the UST Branch any time a leak detector 'trips' and cannot be reset."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        bolAddNone = False
                                    End If

                                    If alElectInter.Contains(drdsFac("FACILITY_ID")) Then

                                        strSummaryDesc = "Electronic Interstitial Monitoring as a method of tank and/or piping leak detection"
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, True)

                                        strSummaryDesc = "Document that the electronic sensors are in communication with the control console (monthly 'sensor status report')."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Maintain a log showing all alarms and the reporting or reconciliation of each alarm."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "The functionality of all electronic interstitial monitoring devices must be tested every 12 months."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Report to the UST Branch any alarm that could indicate a release to the environment has occurred."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        bolAddNone = False
                                    End If

                                    If alVisInter.Contains(drdsFac("FACILITY_ID")) Then

                                        strSummaryDesc = "Visual Interstitial Monitoring as a method of tank and/or piping leak detection"
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, True)

                                        strSummaryDesc = "The interstice must be checked every 30 days and a written record maintained."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Maintain the previous 12 months of records."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Report to the UST Branch whenever you discover any free product in the tank interstice or > 1/8 inch within the pipe sump."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        bolAddNone = False
                                    End If

                                    If alImpress.Contains(drdsFac("FACILITY_ID")) Then

                                        strSummaryDesc = "60 day log of Impressed Current Cathodic Protection system rectifier"
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, True)

                                        strSummaryDesc = "Maintain record that indicates rectifier has been checked for proper operation once every 60 days."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Electrical power to rectifier can not be interrupted except for maintenance activities."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)
                                        strSummaryDesc = "Most manufacturers recommend you investigate if the voltage/amperage output of the rectifier changes by more than 20%."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        bolAddNone = False
                                    End If

                                    If alSO.Contains(drdsFac("FACILITY_ID")) Then

                                        strSummaryDesc = "Spill and Overfill Prevention"
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, True)

                                        strSummaryDesc = "Spill containment must be maintained liquid tight and emptied of fluids prior to and after deliveries."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        '      strSummaryDesc = "Spill containment must be maintained so that they are liquid tight."
                                        '     AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Overfill prevention devices must be maintained in working order and inspected every 12 months."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        ' strSummaryDesc = "Overfill prevention devices must be accessible for inspection."
                                        'AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)
                                        'strSummaryDesc = "Overfill prevention devices must be inspected every 12 months."
                                        'AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Report to the UST Branch any spill or overfill of 25 gallons or more or any size spill/overfill if it reaches the environment."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        bolAddNone = False
                                    End If

                                    If (flagDateSecondaryContainmentLastInspectedSummary) Or (flagDatePipeSecondarySummary) Then
                                        strSummaryDesc = "Secondary Containment of the tank and/or pipe"
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, True)

                                        strSummaryDesc = "Secondary containment must be maintained free of liquids and debris if the interstice is designed to be dry."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Secondary containment integrity must be visually inspected every 12 months unless it is monitored continuously."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)
                                        strSummaryDesc = "Secondary containment must be tested if the visual inspection indicates the integrity may be compromised."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        bolAddNone = False

                                    End If

                                    If alTOSI.Contains(drdsFac("FACILITY_ID")) Then

                                        strSummaryDesc = "Our records indicate that one or more of this facility's tanks are temporarily out of use"
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, True)

                                        strSummaryDesc = "Tanks must be emptied within 90 days of taking out of service."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Vent lines must be left open and functioning."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        strSummaryDesc = "Maintain corrosion protection if the tank system is so equipped."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)

                                        bolAddNone = False
                                    End If

                                    If bolAddNone Then

                                        strSummaryDesc = "There is no routine testing or monitoring required by the UST regulations for this facility." + vbCrLf + _
                                                            "However, you should ensure that the manufacturer of your equipment does not require that any periodic tests be conducted."
                                        AddCAPAnnualSummaryText(drdsOwner, drdsFac, dtSummary, processingYear, nSummaryLineCount, strSummaryDesc, False, False, False)
                                    End If
                                End If
                            End If

                        Next ' facility

                        'End With ' With docSummary

                        ' create calendar
                        ' if bolownerneeds cal = true, create calendar, assessment with cal


                        If mode = CapAnnualMode.StaticByYear Then
                            pOwn.ClearCAPAnnualCalendar(processingYear, drdsOwner("OWNER_ID"), , , fac)
                        End If


                        If bolOwnerNeedsCal And bolAddedSectionForOwner And mode = CapAnnualMode.StaticByYear Then

                            If dtCalInfo.Rows.Count > 0 Then
                                Try

                                    For Each drCal As DataRow In dtCalInfo.Rows

                                        Dim facID As Integer = drCal("FACILITY_ID")
                                        Dim OwnerID As Integer = drdsOwner("OWNER_ID")
                                        Dim OwnerNameStr = drdsOwner("OWNERNAME").ToString.ToUpper
                                        Dim City = drCal("CITY").ToString
                                        Dim facility = drCal("NAME").ToString
                                        Dim requirements = drCal("INFO").ToString
                                        Dim month As Integer = drCal("MONTH")

                                        Try
                                            pOwn.SaveCAPAnnualCalendar(processingYear, _
                                                    OwnerID, _
                                                    OwnerNameStr, _
                                                    month, _
                                                    facID, _
                                                    facility, _
                                                    City, _
                                                    requirements, _
                                                     _container.AppUser.ID)
                                        Catch ex As Exception
                                            If ex.ToString.IndexOf("PK_") = -1 Then
                                                Throw ex
                                            End If
                                        End Try

                                    Next
                                Catch ex As Exception
                                    Throw ex
                                End Try

                                ' AddCAPAnnualCalendar(drdsOwner("OWNERNAME"), dtCalInfo, docCal, bolEnteredTextToCalendar, Not bolEnteredTextToCalendar, processingYear, slMonth)
                                bolEnteredTextToCalendar = True

                                ''AddCAPAnnualAssessment(drdsOwner, docAssessCal, bolEnteredTextToAssessCal, strAssessCalTemplate, processingYear)
                                'AddCAPAnnualAssessment(drdsOwner, docAssessCalDS, processingYear, strUser, strUserPhone, bolEnteredTextToAssessCal)
                                bolEnteredTextToAssessCal = True
                            Else
                                ''AddCAPAnnualAssessment(drdsOwner, docAssessNoCal, bolEnteredTextToAssessNoCal, strAssessNoCalTemplate, processingYear)
                                'AddCAPAnnualAssessment(drdsOwner, docAssessNoCalDS, processingYear, strUser, strUserPhone, bolEnteredTextToAssessNoCal)
                                bolEnteredTextToAssessNoCal = True
                            End If
                        ElseIf Not bolOwnerNeedsCal And bolAddedSectionForOwner Then
                            ''AddCAPAnnualAssessment(drdsOwner, docAssessNoCal, bolEnteredTextToAssessNoCal, strAssessNoCalTemplate, processingYear)
                            ' AddCAPAnnualAssessment(drdsOwner, docAssessNoCalDS, processingYear, strUser, strUserPhone, bolEnteredTextToAssessNoCal)
                            bolEnteredTextToAssessNoCal = True
                        End If

                    End If


                Next ' owner

                'dtEnd = DateTime.Now
                'ts = dtEnd.Subtract(dtStart)
                'strTime += vbCrLf + "generate cap yearly report: " + ts.ToString

                If bolEnteredTextToSummary Then

                    ' if datatable has rows, save to db

                    Dim cntRows As Integer = dtSummary.Rows.Count

                    If cntRows > 0 Then
                        Dim nLinePosition, ownerID, facID As Integer
                        Dim ownName, facName, facAddr1, facCity, facState, facZip, desc, createdBy As String
                        Dim isDescPeriodicTestReq, isDescHeading, isDescSubHeading As Boolean

                        Dim I As Integer = 0
                        Try
                            ' clear any cap annual summary data if present in db
                            If OwnerName Is Nothing OrElse OwnerName.Length = 0 Then
                                pOwn.ClearCAPAnnualSummary(processingYear, mode, 0, fac)
                            Else
                                pOwn.ClearCAPAnnualSummary(processingYear, mode, ownID, fac)

                            End If


                            createdBy = MusterContainer.AppUser.ID

                            Dim modeStr As String

                            If mode = CapAnnualMode.CurrentSummary Then
                                modeStr = "Current"
                            Else
                                modeStr = "Yearly"
                            End If
                            For Each drSummary In dtSummary.Rows


                                If String.Format("{0}   Generating {2} CAP Summary: {1}% ", oldText, Int((((i + 1) / cntRows) * 100)), modeStr) <> _container.Text Then

                                    _container.Text = String.Format("{0}   Generating {2} CAP Summary: {1}% ", oldText, Int((((i + 1) / cntRows) * 100)), modeStr)

                                End If

                                nLinePosition = drSummary("LINE_POSITION")
                                ownerID = drSummary("OWN_ID")
                                facID = drSummary("FAC_ID")
                                If drSummary("OWN_NAME") Is DBNull.Value Then
                                    ownName = String.Empty
                                Else
                                    ownName = drSummary("OWN_NAME")
                                End If
                                If drSummary("FAC_NAME") Is DBNull.Value Then
                                    facName = String.Empty
                                Else
                                    facName = drSummary("FAC_NAME")
                                End If
                                If drSummary("FAC_ADDRESS_LINE_ONE") Is DBNull.Value Then
                                    facAddr1 = String.Empty
                                Else
                                    facAddr1 = drSummary("FAC_ADDRESS_LINE_ONE")
                                End If
                                If drSummary("FAC_CITY") Is DBNull.Value Then
                                    facCity = String.Empty
                                Else
                                    facCity = drSummary("FAC_CITY")
                                End If
                                If drSummary("FAC_STATE") Is DBNull.Value Then
                                    facState = String.Empty
                                Else
                                    facState = drSummary("FAC_STATE")
                                End If
                                If drSummary("FAC_ZIP") Is DBNull.Value Then
                                    facZip = String.Empty
                                Else
                                    facZip = drSummary("FAC_ZIP")
                                End If
                                If drSummary("DESCRIPTION") Is DBNull.Value Then
                                    desc = String.Empty
                                Else
                                    desc = drSummary("DESCRIPTION")
                                End If
                                If drSummary("IS_DESC_PERIODIC_TEST_REQ") Is DBNull.Value Then
                                    isDescPeriodicTestReq = False
                                Else
                                    isDescPeriodicTestReq = drSummary("IS_DESC_PERIODIC_TEST_REQ")
                                End If
                                If drSummary("IS_DESC_HEADING") Is DBNull.Value Then
                                    isDescHeading = False
                                Else
                                    isDescHeading = drSummary("IS_DESC_HEADING")
                                End If
                                If drSummary("IS_DESC_SUB_HEADING") Is DBNull.Value Then
                                    isDescSubHeading = False
                                Else
                                    isDescSubHeading = drSummary("IS_DESC_SUB_HEADING")
                                End If

                                pOwn.SaveCAPAnnualSummary(processingYear, _
                                        nLinePosition, _
                                        ownerID, _
                                        ownName, _
                                        facID, _
                                        facName, _
                                        facAddr1, _
                                        facCity, _
                                        facState, _
                                        facZip, _
                                        desc, _
                                        isDescPeriodicTestReq, _
                                        isDescHeading, _
                                        isDescSubHeading, _
                                        createdBy, mode)

                                I += 1

                            Next
                        Catch ex As Exception

                            _container.Text = oldText


                            Throw ex
                        End Try
                    End If
                End If

                'If Not bolEnteredTextToCalendar Then
                'If Not docCal Is Nothing Then
                'docCal.Close(False)
                'UIUtilsGen.Delay(, 1)
                'System.IO.File.Delete(doc_path + strCalDocName)
                'End If
                'bolDeletedCal = True
                'Else
                ''' Save to print basket
                'UIUtilsGen.SaveDocument(0, 0, strCalDocName, "CAP Annual Calendar", doc_path, "CAP Annual Calendar for " + processingYear.ToString, UIUtilsGen.ModuleID.CAPProcess, 0, 0, 0)
                'End If

                'If Not bolEnteredTextToAssessCal Then
                'If Not docAssessCal Is Nothing Then
                'docAssessCal.Close(False)
                'UIUtilsGen.Delay(, 1)
                'System.IO.File.Delete(doc_path + strAssessCalDocName)
                'End If
                'If Not docAssessCalDS Is Nothing Then
                'docAssessCalDS.Close(False)
                'UIUtilsGen.Delay(, 1)
                'System.IO.File.Delete(strCAPAnnualMailMergeCalDocName)
                'End If
                'bolDeletedAssessCal = True
                'Else
                'MailMergeAnnualAssessment(docAssessCalDS, WordApp, docAssessCal)
                ''' Save to print basket
                'UIUtilsGen.SaveDocument(0, 0, strAssessCalDocName, "CAP Annual Assessment Letter with Calendar", doc_path, "CAP Annual Assessment Letter with Calendar for " + processingYear.ToString, UIUtilsGen.ModuleID.CAPProcess, 0, 0, 0)
                'End If

                'If Not bolEnteredTextToAssessNoCal Then
                'If Not docAssessNoCal Is Nothing Then
                'docAssessNoCal.Close(False)
                'UIUtilsGen.Delay(, 1)
                'System.IO.File.Delete(doc_path + strAssessNoCalDocName)
                'End If
                'If Not docAssessNoCalDS Is Nothing Then
                'docAssessNoCalDS.Close(False)
                'UIUtilsGen.Delay(, 1)
                'System.IO.File.Delete(strCAPAnnualMailMergeNoCalDocName)
                'End If
                'bolDeletedAssessNoCal = True
                'Else
                'MailMergeAnnualAssessment(docAssessNoCalDS, WordApp, docAssessNoCal)
                '''' Save to print basket
                'UIUtilsGen.SaveDocument(0, 0, strAssessNoCalDocName, "CAP Annual Assessment Letter without Calendar", doc_path, "CAP Annual Assessment Letter without Calendar for " + processingYear.ToString, UIUtilsGen.ModuleID.CAPProcess, 0, 0, 0)
                'End If

                'If bolDeletedSummary And bolDeletedCal And bolDeletedAssessCal And bolDeletedAssessNoCal Then

                _container.Text = oldText

                If bolDeletedCal And bolDeletedAssessCal And bolDeletedAssessNoCal Then
                    MsgBox("No Records Found")
                    Return False
                Else
                    Return True
                End If

                ' End With ' with wordapp
            Else
                MsgBox("No Records Found")
                ' delete the docs created at the top as no text was entered to the files
                bolDeleteFilesCreated = True
                Return False
            End If

        Catch ex As Exception

            _container.Text = oldText


            bolDeleteFilesCreated = True
            If ex.Message.StartsWith("Template(") And ex.Message.EndsWith("not found") Then
                MsgBox(ex.Message)
            Else
                Dim MyErr As New ErrorReport(ex)
                MyErr.ShowDialog()
            End If
            Return False
        Finally

            _container.Text = oldText


            'dtEndAll = DateTime.Now
            'ts = dtEndAll.Subtract(dtStartAll)
            'strTime += vbCrLf + "complete process: " + ts.ToString
            'MsgBox(strTime)

            If bolDeleteFilesCreated Then

                '''' delete annual summary docs
                ''If Not docSummary Is Nothing Then
                ''    strSummaryDocName = docSummary.Name
                ''    If Not slCAPAnnualSummaryDocs.Item(0) Is Nothing Then
                ''        slCAPAnnualSummaryDocs.Remove(0)
                ''    End If
                ''    docSummary.Close(False)
                ''    UIUtilsGen.Delay(, 1)
                ''    System.IO.File.Delete(doc_path + strSummaryDocName)
                ''End If
                ''For Each docName As String In slCAPAnnualSummaryDocs.Values
                ''    For Each doc As Word.Document In WordApp.Documents
                ''        If doc.Name = docName Then
                ''            doc.Close(False)
                ''            UIUtilsGen.Delay(, 1)
                ''        End If
                ''    Next
                ''    System.IO.File.Delete(doc_path + docName)
                ''Next

                'If Not docCal Is Nothing Then
                'docCal.Close(False)
                'UIUtilsGen.Delay(, 1)
                'System.IO.File.Delete(doc_path + strCalDocName)
                'End If

                'If Not docAssessCal Is Nothing Then
                'docAssessCal.Close(False)
                'UIUtilsGen.Delay(, 1)
                'System.IO.File.Delete(doc_path + strAssessCalDocName)
                'End If

                'If Not docAssessNoCal Is Nothing Then
                ' docAssessNoCal.Close(False)
                ' UIUtilsGen.Delay(, 1)
                ' System.IO.File.Delete(doc_path + strAssessNoCalDocName)
                'End If

                'If Not docAssessCalDS Is Nothing Then
                'strCAPAnnualMailMergeCalDocName = docAssessCalDS.FullName
                'docAssessCalDS.Close(False)
                'UIUtilsGen.Delay(, 1)
                'System.IO.File.Delete(strCAPAnnualMailMergeCalDocName)
                'End If

                'If Not docAssessNoCalDS Is Nothing Then
                'strCAPAnnualMailMergeNoCalDocName = docAssessNoCalDS.FullName
                'docAssessNoCalDS.Close(False)
                'UIUtilsGen.Delay(, 1)
                'System.IO.File.Delete(strCAPAnnualMailMergeNoCalDocName)
                'End If

            End If
        End Try
    End Function


    Private Sub FormatCalendarTable(ByRef table As Word.Table)


        With table


            .Rows.HeightRule = Word.WdRowHeightRule.wdRowHeightAtLeast
            .Rows.Height = 0.4
            .Range.Font.Name = "Arial"
            .Range.Font.Size = 9



            .Columns.Item(1).PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
            .Columns.Item(1).PreferredWidth = 10

            .Columns.Item(2).PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
            .Columns.Item(2).PreferredWidth = 25
            .Columns.Item(3).PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
            .Columns.Item(3).PreferredWidth = 15
            .Columns.Item(4).PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
            .Columns.Item(4).PreferredWidth = 50

            With .Shading
                .Texture = Word.WdTextureIndex.wdTextureNone
                .ForegroundPatternColor = Word.WdColor.wdColorAutomatic
                .BackgroundPatternColor = Word.WdColor.wdColorAutomatic
            End With

            .Borders.Shadow = False



        End With ' with .tables

    End Sub

    Private Function AddPageBreakToAnnualCalendar(ByVal ownername As String, ByRef doc As Word.Document, ByVal processingYear As Integer, ByVal bolAddSectionBReak As Boolean, ByVal FirstRec As Boolean) As Double

        Dim heightCount As Double = 0

        Exit Function


        With doc

            If bolAddSectionBReak Then
                InsertLines(1, doc)

                ' insert section break
                WordApp.Selection.EndKey(Unit:=Word.WdUnits.wdStory)
                doc.Application.Selection.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage)

            ElseIf Not FirstRec Then
                InsertLines(1, doc)


                ' insert page break
                WordApp.Selection.EndKey(Unit:=Word.WdUnits.wdStory)
                doc.Application.Selection.InsertBreak(Word.WdBreakType.wdPageBreak)


            End If


            ' add table
            .Tables.Add(Range:=.Bookmarks.Item("\endofdoc").Range, NumRows:=1, NumColumns:=3, DefaultTableBehavior:=Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Word.WdAutoFitBehavior.wdAutoFitWindow)

            SetTableBorder(.Tables.Item(.Tables.Count), False)


            With .Tables.Item(.Tables.Count)

                With .Borders.Item(Word.WdBorderType.wdBorderBottom)
                    .Visible = True
                    .LineStyle = Word.WdLineStyle.wdLineStyleSingle
                    .LineWidth = Word.WdLineWidth.wdLineWidth050pt
                    .Color = Word.WdColor.wdColorAutomatic
                End With
                With .Borders.Item(Word.WdBorderType.wdBorderTop)
                    .Visible = True
                    .LineStyle = Word.WdLineStyle.wdLineStyleSingle
                    .LineWidth = Word.WdLineWidth.wdLineWidth050pt
                    .Color = Word.WdColor.wdColorAutomatic
                End With

                .Cell(1, 1).PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                .Cell(1, 1).PreferredWidth = 20
                .Cell(1, 2).PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                .Cell(1, 2).PreferredWidth = 60
                .Cell(1, 3).PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                .Cell(1, 3).PreferredWidth = 20

                With .Shading
                    .Texture = Word.WdTextureIndex.wdTextureNone
                    .ForegroundPatternColor = Word.WdColor.wdColorAutomatic
                    .BackgroundPatternColor = Word.WdColor.wdColorAutomatic
                End With

                .Borders.Shadow = False

            End With


            With .Tables.Item(.Tables.Count)

                .Rows.Item(1).Height = 6

                .Cell(1, 1).Range.InlineShapes.AddPicture(FILENAME:=TmpltPath + "CAP\CAP.gif", _
                    LinkToFile:=False, SaveWithDocument:=True)
                .Cell(1, 3).Range.InlineShapes.AddPicture(FILENAME:=TmpltPath + "CAP\LOGO.gif", _
                    LinkToFile:=False, SaveWithDocument:=True)

                .Cell(1, 2).Select()
                WordApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                .Cell(1, 2).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                .Cell(1, 2).Range.Font.Name = "Arial"
                .Cell(1, 2).Range.Font.Size = 16
                .Cell(1, 2).Range.Font.Bold = 1
                .Cell(1, 2).Range.Text = processingYear.ToString + " Calendar" + vbCrLf + _
                                        "UST TESTING REQUIREMENTS"

                heightCount += 6
            End With ' with .tables

            InsertLines(1, doc)
            heightCount += 0.7

            ' add owner name
            oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
            oPara.Range.Font.Name = "Arial"
            oPara.Range.Font.Size = 14
            oPara.Range.Font.Bold = 1
            oPara.Range.Text = ownername
            oPara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            oPara.Range.InsertParagraphAfter()

            oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
            oPara.Range.Font.Name = "Arial"
            oPara.Range.Font.Size = 5
            oPara.Range.Font.Bold = 1
            oPara.Range.Text = ""
            oPara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            oPara.Range.InsertParagraphAfter()

            heightCount += 3.5

            ' add table
            .Tables.Add(Range:=.Bookmarks.Item("\endofdoc").Range, NumRows:=1, NumColumns:=4, DefaultTableBehavior:=Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Word.WdAutoFitBehavior.wdAutoFitWindow)


            SetTableBorder(.Tables.Item(.Tables.Count), False)

            FormatCalendarTable(.Tables.Item(.Tables.Count))

            With .Tables.Item(.Tables.Count)

                .Rows.Item(1).Shading.BackgroundPatternColor = Word.WdColor.wdColorDarkBlue
                ' heading

                .Cell(1, 1).Range.Text = "Fac #"
                .Cell(1, 1).Range.Font.Size = 12
                .Cell(1, 1).Range.Font.Bold = 1
                .Cell(1, 1).Range.Font.Color = Word.WdColor.wdColorYellow
                .Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight

                .Cell(1, 2).Range.Text = "Facility"
                .Cell(1, 2).Range.Font.Bold = 1
                .Cell(1, 2).Range.Font.Color = Word.WdColor.wdColorYellow
                .Cell(1, 2).Range.Font.Size = 12
                .Cell(1, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft

                .Cell(1, 3).Range.Text = "City"
                .Cell(1, 3).Range.Font.Bold = 1
                .Cell(1, 3).Range.Font.Color = Word.WdColor.wdColorYellow
                .Cell(1, 3).Range.Font.Size = 12
                .Cell(1, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft

                .Cell(1, 4).Range.Text = "Requirements"
                .Cell(1, 4).Range.Font.Bold = 1
                .Cell(1, 4).Range.Font.Color = Word.WdColor.wdColorYellow
                .Cell(1, 4).Range.Font.Size = 12
                .Cell(1, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft


                heightCount += CDbl(.Rows.Item(.Rows.Count).Height)
            End With
        End With


        Return heightCount

    End Function


    Private Function DrawRowForAnnualCalendar(ByVal docTable As Word.Table, ByVal FacID As String, ByVal Facility As String, ByVal City As String, ByVal requirements As String, ByVal cnt As Integer, ByVal printDetails As Boolean) As Double

        Exit Function

        With docTable

            Dim levels As Double = 0.9

            For g As Integer = 1 To 4

                .Cell(cnt, g).Range.Font.Color = Word.WdColor.wdColorBlack
                .Cell(cnt, g).Range.Font.Size = 8
                .Cell(cnt, g).Range.Font.Bold = 0
                .Cell(cnt, g).Shading.BackgroundPatternColor = WdColor.wdColorWhite
                .Cell(cnt, g).Height = levels
                .Cell(cnt, g).HeightRule = WdRowHeightRule.wdRowHeightAuto

                .Cell(cnt, g).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
                .Cell(cnt, g).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft

            Next g

            .Cell(cnt, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight


            If printDetails Then
                .Cell(cnt, 1).Range.Text = FacID
                .Cell(cnt, 2).Range.Text = Facility
                .Cell(cnt, 3).Range.Text = City
            End If

            .Cell(cnt, 4).Range.Text = requirements


            Return levels

        End With

    End Function


    Private Sub MergeCells(ByVal firstCell As Integer, ByVal lastCell As Integer, ByVal docCol As Integer, ByVal doctable As Word.Table, ByVal doc As Word.Document)

        If firstCell <> lastCell Then

            With doctable

                doc.Range(Start:=doctable.Cell(firstCell, docCol).Range.Start, _
                                End:=doctable.Cell(lastCell, docCol).Range.End).Select()



                .Application.Selection.Cells.Merge()

            End With



        End If



    End Sub

    Private Sub AddCAPAnnualCalendar(ByVal ownerName As String, ByVal dtCalInfo As DataTable, ByRef doc As Word.Document, ByVal bolAddSectionBreak As Boolean, ByVal bolFirstRec As Boolean, ByVal processingYear As Integer, ByVal slMonth As SortedList)
        Dim drCal As DataRow
        Dim nMonth As Integer = 0
        Dim facID As String = String.Empty
        Dim City As String = String.Empty
        Dim facility As String = String.Empty
        Dim requirements As String = String.Empty
        Dim x As Integer
        Dim y As Integer
        Dim heightCount As Double = 0
        Dim colParams As New Specialized.NameValueCollection
        Dim strMonth As String = ""
        Dim dataFilled As Boolean = False
        Try

            Exit Sub
            With doc


                doc.Activate()


                heightCount = AddPageBreakToAnnualCalendar(ownerName, doc, processingYear, bolAddSectionBreak, bolFirstRec)


                For i As Integer = 0 To slMonth.Count - 1

                    nMonth = slMonth.GetByIndex(i)

                    Select Case nMonth
                        Case 1
                            strMonth = "January"
                        Case 2
                            strMonth = "February"
                        Case 3
                            strMonth = "March"
                        Case 4
                            strMonth = "April"
                        Case 5
                            strMonth = "May"
                        Case 6
                            strMonth = "June"
                        Case 7
                            strMonth = "July"
                        Case 8
                            strMonth = "August"
                        Case 9
                            strMonth = "September"
                        Case 10
                            strMonth = "October"
                        Case 11
                            strMonth = "November"
                        Case 12
                            strMonth = "December"
                        Case Else
                            strMonth = "Invalid Month"
                    End Select

                    Dim rows As DataRow() = dtCalInfo.Select("MONTH = " + nMonth.ToString)


                    If Not rows Is Nothing AndAlso rows.GetUpperBound(0) > -1 Then


                        If heightCount + 2.3 + (rows.GetUpperBound(0) * 0.7) >= 30 AndAlso dataFilled Then
                            heightCount = Me.AddPageBreakToAnnualCalendar(ownerName, doc, processingYear, False, False)
                        End If

                        ' add table
                        Me.InsertLines(1, doc)

                        .Tables.Add(Range:=doc.Bookmarks.Item("\endofdoc").Range, NumRows:=1, NumColumns:=4, DefaultTableBehavior:=Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Word.WdAutoFitBehavior.wdAutoFitWindow)

                        SetTableBorder(.Tables.Item(.Tables.Count), False)

                        FormatCalendarTable(.Tables.Item(.Tables.Count))

                        With .Tables.Item(.Tables.Count)


                            '.Rows.Add()

                            .Cell(1, 3).Range.Text = strMonth
                            .Cell(1, 3).Range.Font.Color = Word.WdColor.wdColorBlack
                            .Cell(1, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                            .Cell(1, 3).Range.Font.Size = 12
                            .Cell(1, 3).Range.Font.Bold = 1
                            .Rows.Item(1).Height = 0.7

                            heightCount += .Rows.Item(1).Height
                        End With



                        facID = String.Empty
                        City = String.Empty
                        facility = String.Empty
                        requirements = String.Empty


                        Me.InsertLines(1, doc)

                        Dim thisTable As Word.Table

                        thisTable = .Tables.Add(Range:=doc.Bookmarks.Item("\endofdoc").Range, NumRows:=(rows.GetUpperBound(0) + 1), NumColumns:=4, DefaultTableBehavior:=Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Word.WdAutoFitBehavior.wdAutoFitWindow)

                        FormatCalendarTable(thisTable)

                        dataFilled = True

                        facID = String.Empty

                        Dim cnt As Integer = 0

                        x = 1

                        For Each drCal In rows

                            thisTable = doc.Tables.Item(doc.Tables.Count)

                            Dim merge As Boolean = False

                            If facID <> drCal("FACILITY_ID").ToString AndAlso facID <> String.Empty Then
                                merge = True
                            End If

                            facID = drCal("FACILITY_ID").ToString
                            City = drCal("CITY").ToString
                            facility = drCal("NAME").ToString
                            requirements = drCal("INFO").ToString

                            heightCount += DrawRowForAnnualCalendar(thisTable, facID, facility, City, requirements, cnt + 1, merge OrElse cnt = 0)

                            If merge Then

                                y = (cnt + 1) - 1

                                MergeCells(x, y, 1, thisTable, doc)
                                MergeCells(x, y, 2, thisTable, doc)
                                MergeCells(x, y, 3, thisTable, doc)

                                'cnt -= (y - x)

                                'set new point
                                x = (cnt + 1)

                            End If

                            cnt += 1

                        Next

                        thisTable = doc.Tables.Item(doc.Tables.Count)

                        If facID <> String.Empty Then

                            y = (cnt + 1) - 1

                            MergeCells(x, y, 1, thisTable, doc)
                            MergeCells(x, y, 2, thisTable, doc)
                            MergeCells(x, y, 3, thisTable, doc)

                        End If

                    End If

                Next

                Dim loopSave As Boolean = False

                While Not loopSave
                    Try

                        .Save()

                        loopSave = True

                    Catch ex As Exception

                        If Not ex.ToString.ToUpper.IndexOf(" PERMISSION") > -1 Then
                            Throw New Exception(ex.ToString)
                        Else
                            loopSave = False
                            Threading.Thread.Sleep(500)

                        End If

                    End Try
                End While


            End With ' with doc
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub AddCAPAnnualSummaryHeading(ByVal ownerName As String, ByVal drFac As DataRow, ByRef doc As Word.Document, ByVal bolAddSectionBreak As Boolean, ByVal bolAddPageBreak As Boolean, ByRef bolCanSplitSummary As Boolean)
        Try

            Exit Sub

            doc.Activate()
            ' if doc's pages >= 200 create new doc
            If doc.ComputeStatistics(Word.WdStatistic.wdStatisticPages) >= 200 And bolCanSplitSummary Then
                doc.Save()
                Threading.Thread.Sleep(2000)

                Dim strSummaryDocName As String = doc.Name
                strSummaryDocName = strSummaryDocName.Substring(0, strSummaryDocName.Length - 5) + slCAPAnnualSummaryDocs.Count.ToString + ".doc"
                System.IO.File.Copy(DOC_PATH + doc.Name, DOC_PATH + strSummaryDocName)
                slCAPAnnualSummaryDocs.Add(slCAPAnnualSummaryDocs.Count, strSummaryDocName)
                ' clear text in summary doc
                doc.Activate()
                doc.Application.Selection.WholeStory()
                doc.Application.Selection.Delete(Unit:=Word.WdUnits.wdCharacter, Count:=1)
                doc.Save()

                Threading.Thread.Sleep(2000)

                bolAddPageBreak = False
                bolAddSectionBreak = False
            End If
            bolCanSplitSummary = False
            With doc
                doc.Activate()
                If bolAddPageBreak Then
                    ' insert page break
                    doc.Application.Selection.EndKey(Unit:=Word.WdUnits.wdStory)
                    doc.Application.Selection.InsertBreak(Word.WdBreakType.wdPageBreak)
                ElseIf bolAddSectionBreak Then
                    ' insert section break
                    doc.Application.Selection.EndKey(Unit:=Word.WdUnits.wdStory)
                    doc.Application.Selection.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage)
                    InsertLines(1, doc)
                Else
                    InsertLines(1, doc)
                End If

                .Tables.Add(Range:=.Bookmarks.Item("\endofdoc").Range, NumRows:=1, NumColumns:=3, DefaultTableBehavior:=Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Word.WdAutoFitBehavior.wdAutoFitWindow)

                SetTableBorder(.Tables.Item(.Tables.Count), False)

                With .Tables.Item(.Tables.Count)
                    .Cell(1, 1).PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                    .Cell(1, 1).PreferredWidth = 20
                    .Cell(1, 2).PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                    .Cell(1, 2).PreferredWidth = 60
                    .Cell(1, 3).PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                    .Cell(1, 3).PreferredWidth = 20
                    With .Shading
                        .Texture = Word.WdTextureIndex.wdTextureNone
                        .ForegroundPatternColor = Word.WdColor.wdColorAutomatic
                        .BackgroundPatternColor = Word.WdColor.wdColorAutomatic
                    End With

                    .Cell(1, 1).Range.InlineShapes.AddPicture(FILENAME:=TmpltPath + "CAP\CAP.gif", _
                        LinkToFile:=False, SaveWithDocument:=True)
                    .Cell(1, 3).Range.InlineShapes.AddPicture(FILENAME:=TmpltPath + "CAP\LOGO.gif", _
                        LinkToFile:=False, SaveWithDocument:=True)

                    .Cell(1, 2).Range.Font.Name = "Arial"
                    .Cell(1, 2).Range.Font.Size = 11
                    .Cell(1, 2).Range.Font.Bold = 1
                    .Cell(1, 2).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    .Cell(1, 2).Select()
                    WordApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    .Cell(1, 2).Range.Text = ownerName + vbCrLf + _
                                        "FACILITY I.D. #" + drFac("FACILITY_ID").ToString + vbCrLf + _
                                        drFac("NAME").ToString.Trim + vbCrLf + _
                                        drFac("ADDRESS_LINE_ONE").ToString.Trim + vbCrLf + _
                                        drFac("CITY").ToString.Trim + ", " + drFac("STATE").ToString.Trim + " " + drFac("ZIP").ToString.Trim
                End With ' with .tables

                oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                'oPara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                oPara.Range.Font.Name = "Arial"
                oPara.Range.Font.Size = 10
                oPara.Range.Font.Bold = 1
                oPara.Range.Text = "Periodic Testing Requirements"
                oPara.Range.InsertParagraphAfter()

                InsertLines(1, doc)
            End With ' with doc
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub AddCAPAnnualSummaryHeading(ByVal drOwn As DataRow, ByVal drFac As DataRow, ByRef dtSummary As DataTable, ByVal processingYear As Integer, ByRef lineCount As Integer, ByVal bolAddFacility As Boolean)
        Dim drSummary As DataRow
        Try
            If bolAddFacility Then
                '   If dtSummary.Select("OWN_ID = " + drOwn("OWNER_ID").ToString + " AND FAC_ID = " + drFac("FACILITY_ID").ToString).Length = 0 Then
                lineCount += 1
                drSummary = dtSummary.NewRow

                drSummary("PROCESSING_YEAR") = processingYear
                drSummary("LINE_POSITION") = lineCount
                drSummary("OWN_ID") = drOwn("OWNER_ID")
                drSummary("OWN_NAME") = drOwn("OWNERNAME")
                drSummary("FAC_ID") = drFac("FACILITY_ID")
                drSummary("FAC_NAME") = drFac("NAME")
                drSummary("FAC_ADDRESS_LINE_ONE") = drFac("ADDRESS_LINE_ONE")
                drSummary("FAC_CITY") = drFac("CITY")
                drSummary("FAC_STATE") = drFac("STATE")
                drSummary("FAC_ZIP") = drFac("ZIP")
                'drSummary("DESCRIPTION") = DBNull.Value
                'drSummary("IS_DESC_PERIODIC_TEST_REQ") = 0
                'drSummary("IS_DESC_HEADING") = 0
                'drSummary("IS_DESC_SUB_HEADING") = 0

                dtSummary.Rows.Add(drSummary)
                ' End If
            End If

            lineCount += 1
            drSummary = dtSummary.NewRow

            drSummary("PROCESSING_YEAR") = processingYear
            drSummary("LINE_POSITION") = lineCount
            drSummary("OWN_ID") = drOwn("OWNER_ID")
            'drSummary("OWN_NAME") = DBNull.Value
            drSummary("FAC_ID") = drFac("FACILITY_ID")
            'drSummary("NAME") = DBNull.Value
            'drSummary("FAC_ADDRESS_LINE_ONE") = drFac("ADDRESS_LINE_ONE")
            'drSummary("FAC_CITY") = drFac("CITY")
            'drSummary("FAC_STATE") = drFac("STATE")
            'drSummary("FAC_ZIP") = drFac("ZIP")
            drSummary("DESCRIPTION") = "Periodic Inspection and Testing Requirements"
            drSummary("IS_DESC_PERIODIC_TEST_REQ") = 1
            drSummary("IS_DESC_HEADING") = 1
            drSummary("IS_DESC_SUB_HEADING") = 0

            dtSummary.Rows.Add(drSummary)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub AddCAPAnnualSummaryText(ByVal drOwn As DataRow, ByVal drFac As DataRow, ByRef dtSummary As DataTable, ByVal processingYear As Integer, ByRef lineCount As Integer, ByVal strDesc As String, Optional ByVal isDescPeriodicTestReq As Boolean = False, Optional ByVal isDescHeading As Boolean = False, Optional ByVal isDescSubHeading As Boolean = False)
        Dim drSummary As DataRow
        Try
            lineCount += 1
            drSummary = dtSummary.NewRow

            drSummary("PROCESSING_YEAR") = processingYear
            drSummary("LINE_POSITION") = lineCount
            drSummary("OWN_ID") = drOwn("OWNER_ID")
            'drSummary("OWN_NAME") = DBNull.Value
            drSummary("FAC_ID") = drFac("FACILITY_ID")
            'drSummary("NAME") = DBNull.Value
            'drSummary("FAC_ADDRESS_LINE_ONE") = drFac("ADDRESS_LINE_ONE")
            'drSummary("FAC_CITY") = drFac("CITY")
            'drSummary("FAC_STATE") = drFac("STATE")
            'drSummary("FAC_ZIP") = drFac("ZIP")
            drSummary("DESCRIPTION") = strDesc
            drSummary("IS_DESC_PERIODIC_TEST_REQ") = isDescPeriodicTestReq
            drSummary("IS_DESC_HEADING") = isDescHeading
            drSummary("IS_DESC_SUB_HEADING") = isDescSubHeading

            dtSummary.Rows.Add(drSummary)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub AddCAPAnnualAssessment(ByVal drOwner As DataRow, ByRef doc As Word.Document, ByVal bolAddSectionBreak As Boolean, ByVal strAssessTemplate As String, ByVal processingYear As Integer)
        Dim colParams As New Specialized.NameValueCollection
        Dim strKey As String = String.Empty
        Dim strValue As String = String.Empty
        Try
            With doc
                doc.Activate()
                If bolAddSectionBreak Then
                    ' insert section break
                    WordApp.Selection.EndKey(Unit:=Word.WdUnits.wdStory)
                    doc.Application.Selection.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage)

                    ' insert file
                    doc.Application.Selection.InsertFile(FILENAME:=strAssessTemplate, ConfirmConversions:=False, Link:=False, Attachment:=False)

                    ' disable auto spell check
                    doc.Application.Selection.NoProofing = True
                End If

                ' Build NameValueCollection with Tags and Values.
                colParams.Add("<Title>", "Compliance Assessment Letter")
                colParams.Add("<Date>", Format(Now, "MMMM d, yyyy"))

                colParams.Add("<Owner Name>", drOwner("OWNERNAME").ToString.Trim)
                colParams.Add("<Owner Address 1>", drOwner("ADDRESS_LINE_ONE").ToString.Trim)
                If drOwner("ADDRESS_TWO") Is DBNull.Value Then
                    colParams.Add("<Owner Address 2>", drOwner("CITY").ToString.Trim + ", " + drOwner("STATE").ToString.Trim + " " + drOwner("ZIP").ToString.Trim)
                    colParams.Add("<Owner City/State/Zip>", "")
                ElseIf drOwner("ADDRESS_TWO").ToString.Trim = String.Empty Then
                    colParams.Add("<Owner Address 2>", drOwner("CITY").ToString.Trim + ", " + drOwner("STATE").ToString.Trim + " " + drOwner("ZIP").ToString.Trim)
                    colParams.Add("<Owner City/State/Zip>", "")
                Else
                    colParams.Add("<Owner Address 2>", drOwner("ADDRESS_TWO").ToString.Trim)
                    colParams.Add("<Owner City/State/Zip>", drOwner("CITY").ToString.Trim + ", " + drOwner("STATE").ToString.Trim + " " + drOwner("ZIP").ToString.Trim)
                End If

                If drOwner("ORGANIZATION_ID") Is DBNull.Value Then
                    colParams.Add("<Owner Greeting>", drOwner("OWNERNAME").ToString.Trim + ":")
                ElseIf drOwner("ORGANIZATION_ID") = 0 Then
                    colParams.Add("<Owner Greeting>", drOwner("OWNERNAME").ToString.Trim + ":")
                Else
                    colParams.Add("<Owner Greeting>", "Dear " + drOwner("OWNERNAME").ToString.Trim + ":")
                End If

                colParams.Add("<Year>", processingYear.ToString)

                Dim userInfoLocal As MUSTER.Info.UserInfo
                userInfoLocal = MusterContainer.AppUser.RetrieveCAEHead()
                colParams.Add("<User Phone>", CType(userInfoLocal.PhoneNumber, String))
                colParams.Add("<User>", userInfoLocal.Name)

                ' Find and Replace the TAGs with Values.
                For i As Integer = 0 To colParams.Count - 1
                    strKey = colParams.Keys(i).ToString
                    strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                    .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)
                Next
            End With
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub AddCapAnnualAssessment(ByVal drOwner As DataRow, ByVal doc As Word.Document, ByVal processingYear As Integer, ByVal strUser As String, ByVal strUserPhone As String, ByVal addRow As Boolean)
        Dim strGreeting As String
        Try
            If drOwner("ORGANIZATION_ID") Is DBNull.Value Then
                strGreeting = drOwner("OWNERNAME").ToString.Trim + ":"
            ElseIf drOwner("ORGANIZATION_ID") = 0 Then
                strGreeting = drOwner("OWNERNAME").ToString.Trim + ":"
            Else
                strGreeting = "Dear " + drOwner("OWNERNAME").ToString.Trim + ":"
            End If
            With doc.Tables.Item(doc.Tables.Count)
                If addRow Then .Rows.Add()
                FillRow(doc, doc.Tables.Count, .Rows.Count, _
                        Format(Now, "MMMM d, yyyy"), _
                        drOwner("OWNERNAME").ToString.TrimEnd, _
                        drOwner("ADDRESS_LINE_ONE").ToString.TrimEnd, _
                        drOwner("CITY").ToString.Trim + ", " + drOwner("STATE").ToString.Trim + " " + drOwner("ZIP").ToString.TrimEnd, _
                        strGreeting, _
                        strUserPhone, _
                        strUser, _
                        processingYear.ToString)
            End With
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub InsertLines(ByVal LineNum As Integer, ByRef wordApp As Word.Application)
        Dim iCount As Integer
        ' Insert "LineNum" blank lines.
        For iCount = 1 To LineNum
            wordApp.Selection.TypeParagraph()
        Next iCount
    End Sub
    Private Sub InsertLines(ByVal lines As Integer, ByRef doc As Word.Document)
        With doc
            doc.Activate()
            For i As Integer = 1 To lines
                oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                oPara.Range.Text = ""
                oPara.Range.Font.Bold = 0
                oPara.Range.InsertParagraphAfter()
            Next
        End With
    End Sub
    Private Sub SetTableBorder(ByRef tbl As Word.Table, ByVal isVisible As Boolean)
        With tbl
            With .Borders.Item(Word.WdBorderType.wdBorderBottom)
                .Visible = isVisible
            End With
            With .Borders.Item(Word.WdBorderType.wdBorderDiagonalDown)
                .Visible = isVisible
            End With
            With .Borders.Item(Word.WdBorderType.wdBorderDiagonalUp)
                .Visible = isVisible
            End With
            With .Borders.Item(Word.WdBorderType.wdBorderHorizontal)
                .Visible = isVisible
            End With
            With .Borders.Item(Word.WdBorderType.wdBorderLeft)
                .Visible = isVisible
            End With
            With .Borders.Item(Word.WdBorderType.wdBorderRight)
                .Visible = isVisible
            End With
            With .Borders.Item(Word.WdBorderType.wdBorderTop)
                .Visible = isVisible
            End With
            With .Borders.Item(Word.WdBorderType.wdBorderVertical)
                .Visible = isVisible
            End With
        End With
    End Sub

    Private Sub FillRow(ByVal Doc As Word.Document, ByVal table As Integer, ByVal Row As Integer, _
    ByVal strDate As String, ByVal strOwnName As String, ByVal strAdd As String, ByVal strCityStateZip As String, ByVal strGreeting As String, ByVal strUserPhone As String, ByVal strUser As String, ByVal strYear As String)

        With Doc.Tables.Item(table)
            ' Insert the data in the specific cell.
            .Cell(Row, 1).Range.InsertAfter(strDate)
            .Cell(Row, 2).Range.InsertAfter(strOwnName)
            .Cell(Row, 3).Range.InsertAfter(strAdd)
            .Cell(Row, 4).Range.InsertAfter(strCityStateZip)
            .Cell(Row, 5).Range.InsertAfter(strGreeting)
            .Cell(Row, 6).Range.InsertAfter(strUserPhone)
            .Cell(Row, 7).Range.InsertAfter(strUser)
            .Cell(Row, 8).Range.InsertAfter(strYear)
        End With
    End Sub

    Private Function CreateMailMergeDataFileNoCal(ByVal templateDoc As Word.Document, ByVal wrdApp As Word.Application) As Word.Document
        Dim wrdDataDoc As Word.Document
        Dim docCreated As Boolean = False

        strCAPAnnualMailMergeNoCalDocName = "C:\CAPAnnualAssessmentNoCal" + "_" + Today.Month.ToString + "_" + Today.Day.ToString + "_" + Today.Year.ToString + "_" + Today.Hour.ToString + "_" + Today.Minute.ToString + ".doc"
        While docCreated = False
            If System.IO.File.Exists(strCAPAnnualMailMergeNoCalDocName) Then
                Try
                    System.IO.File.Delete(strCAPAnnualMailMergeNoCalDocName)
                    docCreated = True
                Catch ex As Exception
                    strCAPAnnualMailMergeNoCalDocName = "C:\CAPAnnualAssessmentNoCal" + "_" + Today.Month.ToString + "_" + Today.Day.ToString + "_" + Today.Year.ToString + "_" + Today.Minute.ToString + "_" + Today.Second.ToString + ".doc"
                End Try
            Else
                docCreated = True
            End If
        End While

        ' Create a data source at C:\DataDoc.doc containing the field data.
        templateDoc.MailMerge.CreateDataSource(Name:=strCAPAnnualMailMergeNoCalDocName, _
              HeaderRecord:="Date, OwnerName, OwnerAddress, OwnerCityStateZip, OwnerGreeting, UserPhone, User, Year")
        ' Open the file to insert data.
        wrdDataDoc = wrdApp.Documents.Open(strCAPAnnualMailMergeNoCalDocName)
        Return wrdDataDoc
    End Function
    Private Function CreateMailMergeDataFileCal(ByVal templateDoc As Word.Document, ByVal wrdApp As Word.Application) As Word.Document
        Dim wrdDataDoc As Word.Document
        Dim docCreated As Boolean = False

        strCAPAnnualMailMergeCalDocName = "C:\CAPAnnualAssessmentCal" + "_" + Today.Month.ToString + "_" + Today.Day.ToString + "_" + Today.Year.ToString + "_" + Today.Hour.ToString + "_" + Today.Minute.ToString + ".doc"
        While docCreated = False
            If System.IO.File.Exists(strCAPAnnualMailMergeCalDocName) Then
                Try
                    System.IO.File.Delete(strCAPAnnualMailMergeCalDocName)
                    docCreated = True
                Catch ex As Exception
                    strCAPAnnualMailMergeCalDocName = "C:\CAPAnnualAssessmentCal" + "_" + Today.Month.ToString + "_" + Today.Day.ToString + "_" + Today.Year.ToString + "_" + Today.Minute.ToString + "_" + Today.Second.ToString + ".doc"
                End Try
            Else
                docCreated = True
            End If
        End While

        ' Create a data source at C:\DataDoc.doc containing the field data.
        templateDoc.MailMerge.CreateDataSource(Name:=strCAPAnnualMailMergeCalDocName, _
              HeaderRecord:="Date, OwnerName, OwnerAddress, OwnerCityStateZip, OwnerGreeting, UserPhone, User, Year")
        ' Open the file to insert data.
        wrdDataDoc = wrdApp.Documents.Open(strCAPAnnualMailMergeCalDocName)
        Return wrdDataDoc
    End Function

    Private Sub MailMergeAnnualAssessment(ByVal doc As Word.Document, ByVal wordApp As Word.Application, ByVal destDoc As Word.Document)
        Dim wrdSelection As Word.Selection
        Dim wrdMailMerge As Word.MailMerge
        Dim strDocPath As String = doc.FullName
        Dim strDestDocPath As String = destDoc.FullName
        Dim strDestDocPathMM As String
        Try
            doc.Save()
            destDoc.MailMerge.Destination = Word.WdMailMergeDestination.wdSendToNewDocument
            ' Perform mail merge.
            destDoc.MailMerge.Execute(False)
            strDestDocPathMM = wordApp.ActiveDocument.FullName

            ' Close the documents
            destDoc.Saved = True
            destDoc.Close(False)
            System.IO.File.Delete(strDestDocPath)

            For Each wd As Word.Document In wordApp.Documents
                If wd.FullName = strDestDocPathMM Then
                    wd.SaveAs(strDestDocPath)
                    Exit For
                    'wordApp.ActiveDocument.SaveAs(strDestDocPath)
                End If
            Next

            doc.Saved = True
            doc.Close(False)
            System.IO.File.Delete(strDocPath)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub MailMergeAnnualAssessmentCal()
        Dim wrdSelection As Word.Selection
        Dim wrdMailMerge As Word.MailMerge
        'Dim wrdMergeFields As Word.MailMergeFields

        ' Add a new document.
        'wrdDoc = wrdApp.Documents.Open("C:\test.doc")
        'wrdDoc.Select()

        'wrdSelection = wrdApp.Selection()
        'wrdMailMerge = wrdDoc.MailMerge()

        ' Create MailMerge Data file.
        'CreateMailMergeDataFile()

        ' open template

        ' Create a string and insert it in the document.
        'StrToAdd = "State University" & vbCr & _
        '            "Electrical Engineering Department"
        'wrdSelection.ParagraphFormat.Alignment = _
        '            Word.WdParagraphAlignment.wdAlignParagraphCenter
        'wrdSelection.TypeText(StrToAdd)

        'InsertLines(4)

        '' Insert merge data.
        'wrdSelection.ParagraphFormat.Alignment = _
        '            Word.WdParagraphAlignment.wdAlignParagraphLeft
        'wrdMergeFields = wrdMailMerge.Fields()
        'wrdMergeFields.Add(wrdSelection.Range, "FirstName")
        'wrdSelection.TypeText(" ")
        'wrdMergeFields.Add(wrdSelection.Range, "LastName")
        'wrdSelection.TypeParagraph()

        'wrdMergeFields.Add(wrdSelection.Range, "Address")
        'wrdSelection.TypeParagraph()
        'wrdMergeFields.Add(wrdSelection.Range, "CityStateZip")

        'InsertLines(2)

        '' Right justify the line and insert a date field
        '' with the current date.
        'wrdSelection.ParagraphFormat.Alignment = _
        '       Word.WdParagraphAlignment.wdAlignParagraphRight
        'wrdSelection.InsertDateTime( _
        '      DateTimeFormat:="dddd, MMMM dd, yyyy", _
        '      InsertAsField:=False)

        'InsertLines(2)

        '' Justify the rest of the document.
        'wrdSelection.ParagraphFormat.Alignment = _
        '       Word.WdParagraphAlignment.wdAlignParagraphJustify

        'wrdSelection.TypeText("Dear ")
        'wrdMergeFields.Add(wrdSelection.Range, "FirstName")
        'wrdSelection.TypeText(",")
        'InsertLines(2)

        '' Create a string and insert it into the document.
        'StrToAdd = "Thank you for your recent request for next " & _
        '    "semester's class schedule for the Electrical " & _
        '    "Engineering Department. Enclosed with this " & _
        '    "letter is a booklet containing all the classes " & _
        '    "offered next semester at State University.  " & _
        '    "Several new classes will be offered in the " & _
        '    "Electrical Engineering Department next semester.  " & _
        '    "These classes are listed below."
        'wrdSelection.TypeText(StrToAdd)

        'InsertLines(2)

        '' Insert a new table with 9 rows and 4 columns.
        'wrdDoc.Tables.Add(wrdSelection.Range, NumRows:=9, _
        '     NumColumns:=4)

        'With wrdDoc.Tables.Item(1)
        '    ' Set the column widths.
        '    .Columns.Item(1).SetWidth(51, Word.WdRulerStyle.wdAdjustNone)
        '    .Columns.Item(2).SetWidth(170, Word.WdRulerStyle.wdAdjustNone)
        '    .Columns.Item(3).SetWidth(100, Word.WdRulerStyle.wdAdjustNone)
        '    .Columns.Item(4).SetWidth(111, Word.WdRulerStyle.wdAdjustNone)
        '    ' Set the shading on the first row to light gray.
        '    .Rows.Item(1).Cells.Shading.BackgroundPatternColorIndex = _
        '     Word.WdColorIndex.wdGray25
        '    ' Bold the first row.
        '    .Rows.Item(1).Range.Bold = True
        '    ' Center the text in Cell (1,1).
        '    .Cell(1, 1).Range.Paragraphs.Alignment = _
        '              Word.WdParagraphAlignment.wdAlignParagraphCenter

        '    ' Fill each row of the table with data.
        '    FillRow(wrdDoc, 1, "Class Number", "Class Name", "Class Time", _
        '       "Instructor")
        '    FillRow(wrdDoc, 2, "EE220", "Introduction to Electronics II", _
        '              "1:00-2:00 M,W,F", "Dr. Jensen")
        '    FillRow(wrdDoc, 3, "EE230", "Electromagnetic Field Theory I", _
        '              "10:00-11:30 T,T", "Dr. Crump")
        '    FillRow(wrdDoc, 4, "EE300", "Feedback Control Systems", _
        '              "9:00-10:00 M,W,F", "Dr. Murdy")
        '    FillRow(wrdDoc, 5, "EE325", "Advanced Digital Design", _
        '              "9:00-10:30 T,T", "Dr. Alley")
        '    FillRow(wrdDoc, 6, "EE350", "Advanced Communication Systems", _
        '              "9:00-10:30 T,T", "Dr. Taylor")
        '    FillRow(wrdDoc, 7, "EE400", "Advanced Microwave Theory", _
        '              "1:00-2:30 T,T", "Dr. Lee")
        '    FillRow(wrdDoc, 8, "EE450", "Plasma Theory", _
        '              "1:00-2:00 M,W,F", "Dr. Davis")
        '    FillRow(wrdDoc, 9, "EE500", "Principles of VLSI Design", _
        '              "3:00-4:00 M,W,F", "Dr. Ellison")
        'End With

        '' Go to the end of the document.
        'wrdApp.Selection.GoTo(Word.WdGoToItem.wdGoToLine, _
        '           Word.WdGoToDirection.wdGoToLast)

        'InsertLines(2)

        '' Create a string and insert it into the document.
        'StrToAdd = "For additional information regarding the " & _
        '           "Department of Electrical Engineering, " & _
        '           "you can visit our Web site at "
        'wrdSelection.TypeText(StrToAdd)
        '' Insert a hyperlink to the Web page.
        'wrdSelection.Hyperlinks.Add(Anchor:=wrdSelection.Range, _
        '   Address:="http://www.ee.stateu.tld")
        '' Create a string and insert it in the document.
        'StrToAdd = ".  Thank you for your interest in the classes " & _
        '           "offered in the Department of Electrical " & _
        '           "Engineering.  If you have any other questions, " & _
        '           "please feel free to give us a call at " & _
        '           "555-1212." & vbCr & vbCr & _
        '           "Sincerely," & vbCr & vbCr & _
        '           "Kathryn M. Hinsch" & vbCr & _
        '           "Department of Electrical Engineering" & vbCr
        'wrdSelection.TypeText(StrToAdd)

        ' Perform mail merge.
        'wrdMailMerge.Destination = Word.WdMailMergeDestination.wdSendToNewDocument
        wrdMailMerge.Execute(False)

        ' Close the original form document.
        'wrdDoc.Saved = True
        'wrdDoc.Close(False)
        System.IO.File.Delete("C:\test.doc")
        'wrdApp.ActiveDocument.SaveAs("c:\test.doc")

        ' Release References.
        'wrdSelection = Nothing
        'wrdMailMerge = Nothing
        'wrdMergeFields = Nothing
        'wrdDoc = Nothing
        'wrdApp = Nothing

        ' Clean up temp file.
        System.IO.File.Delete("C:\DataDoc.doc")
    End Sub






    Friend Sub GenerateCAPMonthlyForAll(ByVal processingMonthYear As Date, ByVal pOwn As MUSTER.BusinessLogic.pOwner, Optional ByVal ownerName As String = "")
        Dim ds As DataSet
        Dim dsRelOwnerFac, dsRelFacTank, dsRelTankPipe As DataRelation
        Dim dtOwner, dtFac, dtTank, dtPipe As DataTable
        Dim drOwner, drFac, drTank, drPipe As DataRow ' using different variables for less confusion
        Dim drdsOwner, drdsFac, drdsTank, drdsPipe As DataRow ' for looping through the tables
        Dim drdsFacs() As DataRow
        Dim dtProcessingStart, dtProcessingEnd As Date
        Dim strTestReqDocName As String = ""
        Dim strAssistDocName As String = ""
        Dim strTemplate As String = ""
        Dim strAssistTemplate As String = ""

        Dim ListUnknown As String = String.Empty


        Dim bolDeleteFilesCreated As Boolean = False

        Dim headingText As String = ""
        Dim dt, dt1 As Date

        Dim bolEnteredTextToTestReqLetter As Boolean = False
        Dim bolEnteredTextToAssistLetter As Boolean = False

        Dim bolAddedSectionForOwner As Boolean = False
        Dim bolAddedFacilityDetail As Boolean = False

        Dim docAssist, docTestReq As Word.Document

        Dim alTnkLastTCPDate As New ArrayList
        Dim alTnkLinedDate As New ArrayList
        Dim alTnkTTDate As New ArrayList
        Dim alTnkICExpiresDate As New ArrayList
        Dim alTnkSpillDate As New ArrayList
        Dim alTnkOverfillDate As New ArrayList
        Dim alTnkSecondaryDate As New ArrayList
        Dim alTnkElectronicDate As New ArrayList
        Dim alTnkATGDate As New ArrayList

        Dim alPipeCPDate As New ArrayList
        Dim alPipeALLDDate As New ArrayList
        Dim alPipeTermCPDate As New ArrayList
        Dim alPipeLineUSDate As New ArrayList
        Dim alPipeLinePressDate As New ArrayList
        Dim alPipeSheerDate As New ArrayList
        Dim alPipeSecondaryDate As New ArrayList
        Dim alPipeElectronicDate As New ArrayList
        Dim strOwnerIDs As String = String.Empty

        Try


            If DOC_PATH = "\" Then
                MsgBox("Document Path Unspecified. Please give the path before generating the letter.")
                Exit Sub
            End If

            WordApp = MusterContainer.GetWordApp

            If Not WordApp Is Nothing Then

                If Date.Compare(processingMonthYear, CDate("01/01/0001")) = 0 Then
                    processingMonthYear = CDate(Today.Month.ToString + "/1/" + Today.Year.ToString)
                ElseIf processingMonthYear.Day <> 1 Then
                    processingMonthYear = CDate(processingMonthYear.Month.ToString + "/1/" + processingMonthYear.Year.ToString)
                End If

                ' dtProcessingstart = 2 months from processingMonthYear
                dtProcessingStart = DateAdd(DateInterval.Month, 2, processingMonthYear)
                ' add a month and substract a day to get the last day of the month
                dtProcessingEnd = DateAdd(DateInterval.Month, 1, dtProcessingStart)
                dtProcessingEnd = DateAdd(DateInterval.Day, -1, dtProcessingEnd)

                'To avoid duplicate creation of Letters.
                strTestReqDocName = "REG_CAP_PROCESSING_RPT_" + processingMonthYear.Month.ToString + "-" + processingMonthYear.Year.ToString + ownerName + ".doc"
                strAssistDocName = "REG_Compliance_Assistance_Letter_" + processingMonthYear.Month.ToString + "-" + processingMonthYear.Year.ToString + ownerName + ".doc"

                If FileExists(DOC_PATH + strTestReqDocName) OrElse FileExists(DOC_PATH + strAssistDocName) Then
                    If MsgBox("CAP Processing Report for " + IIf(ownerName.Length > 0, String.Format(" Owner '{0}' for", ownerName), String.Empty) + processingMonthYear.Month.ToString + "-" + processingMonthYear.Year.ToString + "  has been created already. Would like to regenerate this report? ", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then

                        File.Delete(DOC_PATH + strTestReqDocName)
                        File.Delete(DOC_PATH + strAssistDocName)

                        Threading.Thread.Sleep(2000)

                    Else
                        Exit Sub
                    End If

                End If


                'dtEnd = DateTime.Now
                'ts = dtEnd.Subtract(dtStart)
                'strTime += "check path and processing month/year: " + ts.ToString

                'dtStart = DateTime.Now

                ds = pOwn.RunSQLQuery("EXEC spSelCapProcessForAll 0, '" + dtProcessingStart.ToShortDateString + "', '" + dtProcessingEnd.ToShortDateString + "'")


                'dtEnd = DateTime.Now
                'ts = dtEnd.Subtract(dtStart)
                'strTime += vbCrLf + "get results from db: " + ts.ToString

                If ds.Tables(0).Rows.Count > 0 Then ' if owner has no rows, facility / tanks / pipes will not have rows

                    'dtStart = DateTime.Now

                    ' create file for Test Req
                    strTemplate = TmpltPath + "CAP\CapMonthlyTestingReqHeading.doc"
                    If Not System.IO.File.Exists(strTemplate) Then
                        MsgBox("Template(" + strTemplate + " not found")
                        Exit Sub
                    End If

                    strTemplate = TmpltPath + "CAP\CapMonthlyTestReq.doc"
                    If Not System.IO.File.Exists(strTemplate) Then
                        MsgBox("Template(" + strTemplate + " not found")
                        Exit Sub
                    End If
                    System.IO.File.Copy(strTemplate, doc_path + strTestReqDocName)

                    ' create file for Assist
                    strAssistTemplate = TmpltPath + "CAP\CapMonthlyAssistance.doc"
                    If Not System.IO.File.Exists(strAssistTemplate) Then
                        MsgBox("Template(" + strAssistTemplate + " not found")
                        Exit Sub
                    End If
                    System.IO.File.Copy(strAssistTemplate, doc_path + strAssistDocName)

                    'dtEnd = DateTime.Now
                    'ts = dtEnd.Subtract(dtStart)
                    'strTime += vbCrLf + "copy templates to temp folder: " + ts.ToString

                    'dtStart = DateTime.Now


                    'dtEnd = DateTime.Now
                    'ts = dtEnd.Subtract(dtStart)
                    'strTime += vbCrLf + "get wordapp: " + ts.ToString

                    'dtStart = DateTime.Now

                    docAssist = WordApp.Documents.Open(doc_path + strAssistDocName)
                    docTestReq = WordApp.Documents.Open(doc_path + strTestReqDocName)

                    ' enter date in footer
                    docTestReq.Activate()
                    With docTestReq
                        .ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter
                        .Application.Selection.Find.Execute(FindText:="<Date>", ReplaceWith:=Now.Date.ToShortDateString, Replace:=Word.WdReplace.wdReplaceAll)
                        .ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument
                    End With

                    'dtEnd = DateTime.Now
                    'ts = dtEnd.Subtract(dtStart)
                    'strTime += vbCrLf + "open word documents: " + ts.ToString

                    'dtStart = DateTime.Now

                    headingText = GetCapMonthlyNoticeOfTestReqHeading(WordApp)

                    'dtEnd = DateTime.Now
                    'ts = dtEnd.Subtract(dtStart)
                    'strTime += vbCrLf + "get monthly notice of test req heading from word doc: " + ts.ToString

                    'dsRelOwnerFac = New DataRelation("OwnerToFacility", ds.Tables(0).Columns("OWNER_ID"), ds.Tables(1).Columns("OWNER_ID"), False)
                    'dsRelFacTank = New DataRelation("FacilityToTank", ds.Tables(1).Columns("FACILITY_ID"), ds.Tables(2).Columns("FACILITY_ID"), False)
                    'dsRelTankPipe = New DataRelation("TankToPipe", ds.Tables(2).Columns("TANK ID"), ds.Tables(3).Columns("TANK ID"), False)
                    'ds.Relations.Add(dsRelOwnerFac)
                    'ds.Relations.Add(dsRelFacTank)
                    'ds.Relations.Add(dsRelTankPipe)

                    'ug.DataSource = ds

                    ' datatables to maintain the records of the tanks and pipes whose dates are to be rolledover
                    dtOwner = New DataTable
                    dtFac = New DataTable
                    dtTank = New DataTable
                    dtPipe = New DataTable

                    dtOwner.Columns.Add("OWNER_ID", GetType(Integer))

                    dtFac.Columns.Add("OWNER_ID", GetType(Integer))
                    dtFac.Columns.Add("FACILITY_ID", GetType(Integer))

                    dtTank.Columns.Add("FACILITY_ID", GetType(Integer))
                    dtTank.Columns.Add("TANK ID", GetType(Integer))
                    dtTank.Columns.Add("CP DATE", GetType(Date))
                    dtTank.Columns.Add("LI INSPECTED", GetType(Date))
                    dtTank.Columns.Add("TT DATE", GetType(Date))

                    'added by Hua Cao 11/12/2008
                    dtTank.Columns.Add("DateSpillPreventionInstalled", GetType(Date))
                    dtTank.Columns.Add("DateSpillPreventionLastTested", GetType(Date))
                    dtTank.Columns.Add("DateOverfillPreventionLastInspected", GetType(Date))
                    dtTank.Columns.Add("DateSecondaryContainmentLastInspected", GetType(Date))
                    dtTank.Columns.Add("DateElectronicDeviceInspected", GetType(Date))
                    dtTank.Columns.Add("DateATGLastInspected", GetType(Date))
                    dtTank.Columns.Add("DateOverfillPreventionInstalled", GetType(Date))


                    dtPipe.Columns.Add("TANK ID", GetType(Integer))
                    dtPipe.Columns.Add("PIPE ID", GetType(Integer))
                    dtPipe.Columns.Add("CP DATE", GetType(Date))
                    dtPipe.Columns.Add("TERM CP TEST", GetType(Date))
                    dtPipe.Columns.Add("ALLD_TEST_DATE", GetType(Date))
                    dtPipe.Columns.Add("TT DATE", GetType(Date))
                    'added by Hua Cao 11/12/2008
                    dtPipe.Columns.Add("DateSheerValueTest", GetType(Date))
                    dtPipe.Columns.Add("DateSecondaryContainmentInspect", GetType(Date))
                    dtPipe.Columns.Add("DateElectronicDeviceInspect", GetType(Date))

                    With WordApp

                        .Visible = True

                        'dtStart = DateTime.Now
                        'Owner loop for CAP monthly report


                        For i As Integer = 0 To ds.Tables(0).Rows.Count - 1 ' owner
                            'For Each drdsOwner In ds.Tables(0).Rows ' owner
                            drdsOwner = ds.Tables(0).Rows(i)

                            If ownerName = String.Empty OrElse drdsOwner("OWNERNAME").ToString.ToUpper.IndexOf(ownerName.ToUpper) > -1 Then

                                drOwner = dtOwner.NewRow
                                drOwner("OWNER_ID") = drdsOwner("OWNER_ID")

                                With docAssist

                                    bolAddedSectionForOwner = False

                                    drdsFacs = ds.Tables(1).Select("OWNER_ID = " + drdsOwner("OWNER_ID").ToString)


                                    Dim points As Integer = 0

                                    'Facility Loop for CAP monthly report
                                    For j As Integer = 0 To drdsFacs.Length - 1 ' facility
                                        'For Each drdsFac In ds.Tables(1).Select("OWNER_ID = " + drdsOwner("OWNER_ID").ToString) ' facility
                                        drdsFac = drdsFacs(j)

                                        drFac = dtFac.NewRow
                                        drFac("OWNER_ID") = drdsOwner("OWNER_ID")
                                        drFac("FACILITY_ID") = drdsFac("FACILITY_ID")

                                        bolAddedFacilityDetail = False

                                        alTnkLastTCPDate = New ArrayList
                                        alTnkLinedDate = New ArrayList
                                        alTnkTTDate = New ArrayList
                                        alTnkICExpiresDate = New ArrayList
                                        alTnkSpillDate = New ArrayList
                                        alTnkOverfillDate = New ArrayList
                                        alTnkSecondaryDate = New ArrayList
                                        alTnkElectronicDate = New ArrayList
                                        alTnkATGDate = New ArrayList

                                        alPipeCPDate = New ArrayList
                                        alPipeALLDDate = New ArrayList
                                        alPipeTermCPDate = New ArrayList
                                        alPipeLineUSDate = New ArrayList
                                        alPipeLinePressDate = New ArrayList
                                        alPipeSheerDate = New ArrayList
                                        alPipeSecondaryDate = New ArrayList
                                        alPipeElectronicDate = New ArrayList

                                        With docTestReq
                                            docTestReq.Activate()
                                            'Tank loop for the CAP monthly report

                                            For Each drdsTank In ds.Tables(2).Select("FACILITY_ID = " + drdsFac("FACILITY_ID").ToString) ' tank
                                                If drdsFac("FACILITY_ID").ToString = "3161" Then
                                                    Dim test As String
                                                    test = "0"
                                                End If
                                                drTank = dtTank.NewRow
                                                drTank("FACILITY_ID") = drdsFac("FACILITY_ID")
                                                drTank("TANK ID") = drdsTank("TANK ID")

                                                ' check tank conditions
                                                ' if tank modified, set tank's, facility's and owner's MODIFIED column value to true
                                                ' and update the field modified
                                                'added 
                                                'DateSpillPreventionLastTested
                                                'DateOverfillPreventionLastInspected
                                                'DateSecondaryContainmentLastInspected
                                                'DateElectronicDeviceInspected
                                                'DateATGLastInspected

                                                If Not drdsTank("STATUS") Is DBNull.Value Then
                                                    If drdsTank("STATUS").ToString.IndexOf("Currently In Use") > -1 Or drdsTank("STATUS").ToString.IndexOf("Temporarily Out of Service Indefinitely") > -1 Then

                                                        If drdsTank("FACILITY_ID") = 4612 Then
                                                            dt = Nothing

                                                        End If

                                                        'DateSpillPreventionLastTested
                                                        If drdsTank("DateSpillPreventionLastTested") Is DBNull.Value Then
                                                            ListUnknown = "'Unknown'"
                                                            dt = New Date(2009, 10, 1)
                                                        Else
                                                            dt = drdsTank("DateSpillPreventionLastTested")
                                                            dt = dt.Date
                                                            dt = DateAdd(DateInterval.Year, 1, dt)
                                                            ListUnknown = dt.ToShortDateString

                                                        End If


                                                        If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                            If Not alTnkSpillDate.Contains(dt.ToShortDateString) Then
                                                                If Not bolAddedSectionForOwner Then
                                                                    AddCAPMonthlyNoticeOfTestReqHeading(headingText, drdsOwner("OWNERNAME").ToString, docTestReq, bolEnteredTextToTestReqLetter)
                                                                    bolAddedSectionForOwner = True
                                                                End If
                                                                If Not bolAddedFacilityDetail Then
                                                                    points = AddCapMontlyNoticeOfTestReqFacilityDetails(drdsFac, docTestReq, points, headingText, drdsOwner("OWNERNAME").ToString)
                                                                    bolAddedFacilityDetail = True
                                                                End If
                                                                alTnkSpillDate.Add(dt.ToShortDateString)

                                                                points = Me.AddCapMonhlyPageBreak(points + 2, points, docTestReq, headingText, drdsOwner("OWNERNAME").ToString)

                                                                oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                                                                oPara.Range.Font.Name = "Arial"
                                                                oPara.Range.Font.Size = 10
                                                                oPara.Range.Font.Bold = 0
                                                                oPara.Range.Text = "Testing of spill containment buckets must be accomplished by " + ListUnknown + "." + vbCrLf + _
                                                                                    "Please update my records to reflect that this test was accomplished on _______________."
                                                                oPara.Range.InsertParagraphAfter()
                                                                InsertLines(1, docTestReq)

                                                                points += 3
                                                            End If

                                                            bolEnteredTextToTestReqLetter = True
                                                            drdsTank("MODIFIED") = True
                                                            drdsFac("MODIFIED") = True
                                                            drdsOwner("MODIFIED") = True
                                                            drdsTank("DateSpillPreventionLastTested") = dt
                                                            drTank("DateSpillPreventionLastTested") = dt
                                                        End If

                                                        'DateOverfillPreventionLastInspected
                                                        If Not drdsTank("DateOverfillPreventionLastInspected") Is DBNull.Value Then
                                                            dt = drdsTank("DateOverfillPreventionLastInspected")
                                                            dt = dt.Date
                                                            dt = DateAdd(DateInterval.Year, 1, dt)
                                                            ListUnknown = dt.ToShortDateString
                                                        Else
                                                            ListUnknown = "'Unknown'"
                                                            dt = New Date(2009, 10, 1)

                                                        End If

                                                        If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                            If Not alTnkOverfillDate.Contains(dt.ToShortDateString) Then
                                                                If Not bolAddedSectionForOwner Then
                                                                    AddCAPMonthlyNoticeOfTestReqHeading(headingText, drdsOwner("OWNERNAME").ToString, docTestReq, bolEnteredTextToTestReqLetter)
                                                                    bolAddedSectionForOwner = True
                                                                End If
                                                                If Not bolAddedFacilityDetail Then
                                                                    points = AddCapMontlyNoticeOfTestReqFacilityDetails(drdsFac, docTestReq, points, headingText, drdsOwner("OWNERNAME").ToString)
                                                                    bolAddedFacilityDetail = True
                                                                End If
                                                                alTnkOverfillDate.Add(dt.ToShortDateString)

                                                                points = Me.AddCapMonhlyPageBreak(points + 2, points, docTestReq, headingText, drdsOwner("OWNERNAME").ToString)

                                                                oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)

                                                                oPara.Range.Font.Name = "Arial"
                                                                oPara.Range.Font.Size = 10
                                                                oPara.Range.Font.Bold = 0
                                                                oPara.Range.Text = "Inspection of overfill prevention must be accomplished by " + ListUnknown + "." + vbCrLf + _
                                                                                    "Please update my records to reflect that this test was accomplished on _______________."
                                                                oPara.Range.InsertParagraphAfter()
                                                                InsertLines(1, docTestReq)

                                                                points += 3
                                                            End If
                                                            bolEnteredTextToTestReqLetter = True
                                                            drdsTank("MODIFIED") = True
                                                            drdsFac("MODIFIED") = True
                                                            drdsOwner("MODIFIED") = True
                                                            drdsTank("DateOverfillPreventionLastInspected") = dt
                                                            drTank("DateOverfillPreventionLastInspected") = dt
                                                        End If

                                                        'DateSecondaryContainmentLastInspected
                                                        If drdsTank("Tank_LD_Num") = 343 Then
                                                            If Not drdsTank("DateSecondaryContainmentLastInspected") Is DBNull.Value Then
                                                                dt = drdsTank("DateSecondaryContainmentLastInspected")
                                                                dt = dt.Date
                                                                dt = DateAdd(DateInterval.Year, 1, dt)
                                                                ListUnknown = dt.ToShortDateString

                                                            Else
                                                                ListUnknown = "'Unknown'"
                                                                dt = New Date(2009, 10, 1)

                                                            End If

                                                            If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                                If Not alTnkSecondaryDate.Contains(dt.ToShortDateString) Then
                                                                    If Not bolAddedSectionForOwner Then
                                                                        AddCAPMonthlyNoticeOfTestReqHeading(headingText, drdsOwner("OWNERNAME").ToString, docTestReq, bolEnteredTextToTestReqLetter)
                                                                        bolAddedSectionForOwner = True
                                                                    End If
                                                                    If Not bolAddedFacilityDetail Then
                                                                        points = AddCapMontlyNoticeOfTestReqFacilityDetails(drdsFac, docTestReq, points, headingText, drdsOwner("OWNERNAME").ToString)
                                                                        bolAddedFacilityDetail = True
                                                                    End If
                                                                    alTnkSecondaryDate.Add(dt.ToShortDateString)

                                                                    points = Me.AddCapMonhlyPageBreak(points + 2, points, docTestReq, headingText, drdsOwner("OWNERNAME").ToString)

                                                                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                                                                    oPara.Range.Font.Name = "Arial"
                                                                    oPara.Range.Font.Size = 10
                                                                    oPara.Range.Font.Bold = 0
                                                                    oPara.Range.Text = "Inspection of the tank secondary containment must be accomplished by " + ListUnknown + "." + vbCrLf + _
                                                                                        "Please update my records to reflect that this test was accomplished on _______________."
                                                                    oPara.Range.InsertParagraphAfter()
                                                                    InsertLines(1, docTestReq)

                                                                    points += 3
                                                                End If
                                                                bolEnteredTextToTestReqLetter = True
                                                                drdsTank("MODIFIED") = True
                                                                drdsFac("MODIFIED") = True
                                                                drdsOwner("MODIFIED") = True
                                                                drdsTank("DateSecondaryContainmentLastInspected") = dt
                                                                drTank("DateSecondaryContainmentLastInspected") = dt
                                                            End If
                                                        End If


                                                        'DateElectronicDeviceInspected

                                                        If drdsTank("Tank_LD_Num") = 339 Then

                                                            If Not drdsTank("DateElectronicDeviceInspected") Is DBNull.Value Then
                                                                dt = drdsTank("DateElectronicDeviceInspected")
                                                                dt = dt.Date
                                                                dt = DateAdd(DateInterval.Year, 1, dt)
                                                                ListUnknown = dt.ToShortDateString

                                                            Else
                                                                ListUnknown = "'Unknown'"
                                                                dt = New Date(2009, 10, 1)
                                                            End If

                                                            If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                                If Not alTnkElectronicDate.Contains(dt.ToShortDateString) Then
                                                                    If Not bolAddedSectionForOwner Then
                                                                        AddCAPMonthlyNoticeOfTestReqHeading(headingText, drdsOwner("OWNERNAME").ToString, docTestReq, bolEnteredTextToTestReqLetter)
                                                                        bolAddedSectionForOwner = True
                                                                    End If
                                                                    If Not bolAddedFacilityDetail Then
                                                                        points = AddCapMontlyNoticeOfTestReqFacilityDetails(drdsFac, docTestReq, points, headingText, drdsOwner("OWNERNAME").ToString)
                                                                        bolAddedFacilityDetail = True
                                                                    End If
                                                                    alTnkElectronicDate.Add(dt.ToShortDateString)

                                                                    points = Me.AddCapMonhlyPageBreak(points + 2, points, docTestReq, headingText, drdsOwner("OWNERNAME").ToString)

                                                                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                                                                    oPara.Range.Font.Name = "Arial"
                                                                    oPara.Range.Font.Size = 10
                                                                    oPara.Range.Font.Bold = 0
                                                                    oPara.Range.Text = "Testing of tank electronic interstitial monitoring devices must be accomplished by " + ListUnknown + "." + vbCrLf + _
                                                                                        "Please update my records to reflect that this test was accomplished on _______________."
                                                                    oPara.Range.InsertParagraphAfter()
                                                                    InsertLines(1, docTestReq)

                                                                    points += 3
                                                                End If
                                                                bolEnteredTextToTestReqLetter = True
                                                                drdsTank("MODIFIED") = True
                                                                drdsFac("MODIFIED") = True
                                                                drdsOwner("MODIFIED") = True
                                                                drdsTank("DateElectronicDeviceInspected") = dt
                                                                drTank("DateElectronicDeviceInspected") = dt
                                                            End If
                                                        End If


                                                        'DateATGLastInspected
                                                        If drdsTank("Tank_LD_Num") = 336 Then


                                                            If Not drdsTank("DateATGLastInspected") Is DBNull.Value Then

                                                                dt = drdsTank("DateATGLastInspected")
                                                                dt = dt.Date
                                                                dt = DateAdd(DateInterval.Year, 1, dt)
                                                                ListUnknown = dt.ToShortDateString

                                                            Else
                                                                ListUnknown = "'Unknown'"
                                                                dt = New Date(2009, 10, 1)
                                                            End If

                                                            If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                                If Not alTnkATGDate.Contains(dt.ToShortDateString) Then
                                                                    If Not bolAddedSectionForOwner Then
                                                                        AddCAPMonthlyNoticeOfTestReqHeading(headingText, drdsOwner("OWNERNAME").ToString, docTestReq, bolEnteredTextToTestReqLetter)
                                                                        bolAddedSectionForOwner = True
                                                                    End If
                                                                    If Not bolAddedFacilityDetail Then
                                                                        points = AddCapMontlyNoticeOfTestReqFacilityDetails(drdsFac, docTestReq, points, headingText, drdsOwner("OWNERNAME").ToString)
                                                                        bolAddedFacilityDetail = True
                                                                    End If
                                                                    alTnkATGDate.Add(dt.ToShortDateString)

                                                                    points = Me.AddCapMonhlyPageBreak(points + 2, points, docTestReq, headingText, drdsOwner("OWNERNAME").ToString)

                                                                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                                                                    oPara.Range.Font.Name = "Arial"
                                                                    oPara.Range.Font.Size = 10
                                                                    oPara.Range.Font.Bold = 0
                                                                    oPara.Range.Text = "Inspection of automatic tank gauging equipment must be accomplished by " + ListUnknown + "." + vbCrLf + _
                                                                                        "Please update my records to reflect that this test was accomplished on _______________."
                                                                    oPara.Range.InsertParagraphAfter()
                                                                    InsertLines(1, docTestReq)

                                                                    points += 3
                                                                End If
                                                                bolEnteredTextToTestReqLetter = True
                                                                drdsTank("MODIFIED") = True
                                                                drdsFac("MODIFIED") = True
                                                                drdsOwner("MODIFIED") = True
                                                                drdsTank("DateATGLastInspected") = dt
                                                                drTank("DateATGLastInspected") = dt
                                                            End If

                                                        End If


                                                        ' LAST TCP DATE
                                                        If Not drdsTank("TANKMODDESC") Is DBNull.Value Then

                                                            If drdsTank("TANKMODDESC").ToString.IndexOf("Cathodically Protected") > -1 Then
                                                                If Not drdsTank("CP DATE") Is DBNull.Value Then
                                                                    dt = drdsTank("CP DATE")
                                                                    dt = dt.Date
                                                                    dt = DateAdd(DateInterval.Year, 3, dt)
                                                                    ListUnknown = dt.ToShortDateString

                                                                Else
                                                                    ListUnknown = "'Unknown'"
                                                                    dt = New Date(2009, 10, 1)
                                                                End If

                                                                If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                                    If Not alTnkLastTCPDate.Contains(dt.ToShortDateString) Then
                                                                        If Not bolAddedSectionForOwner Then
                                                                            AddCAPMonthlyNoticeOfTestReqHeading(headingText, drdsOwner("OWNERNAME").ToString, docTestReq, bolEnteredTextToTestReqLetter)
                                                                            bolAddedSectionForOwner = True
                                                                        End If
                                                                        If Not bolAddedFacilityDetail Then
                                                                            points = AddCapMontlyNoticeOfTestReqFacilityDetails(drdsFac, docTestReq, points, headingText, drdsOwner("OWNERNAME").ToString)
                                                                            bolAddedFacilityDetail = True
                                                                        End If
                                                                        alTnkLastTCPDate.Add(dt.ToShortDateString)

                                                                        points = Me.AddCapMonhlyPageBreak(points + 2, points, docTestReq, headingText, drdsOwner("OWNERNAME").ToString)


                                                                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                                                                        'oPara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                                                                        oPara.Range.Font.Name = "Arial"
                                                                        oPara.Range.Font.Size = 10
                                                                        oPara.Range.Font.Bold = 0
                                                                        oPara.Range.Text = "Testing of the tank cathodic protection must be accomplished by " + ListUnknown + "." + vbCrLf + _
                                                                                            "Please update my records to reflect that this test was accomplished on _______________."
                                                                        oPara.Range.InsertParagraphAfter()

                                                                        InsertLines(1, docTestReq)

                                                                        points += 3
                                                                    End If

                                                                    bolEnteredTextToTestReqLetter = True
                                                                    drdsTank("MODIFIED") = True
                                                                    drdsFac("MODIFIED") = True
                                                                    drdsOwner("MODIFIED") = True
                                                                    drdsTank("CP DATE") = dt
                                                                    drTank("CP DATE") = dt

                                                                End If
                                                            End If
                                                        End If

                                                        ' LINED DUE
                                                        If Not drdsTank("TANKMODDESC") Is DBNull.Value Then
                                                            If drdsTank("TANKMODDESC").ToString.IndexOf("Lined Interior") > -1 Then
                                                                'enableTankLIInspectedDate = True
                                                                dt = IIf(drdsTank("LI INSTALL") Is DBNull.Value, CDate("01/01/0001"), drdsTank("LI INSTALL"))
                                                                dt1 = IIf(drdsTank("LI INSPECTED") Is DBNull.Value, CDate("01/01/0001"), drdsTank("LI INSPECTED"))
                                                                dt = dt.Date
                                                                dt1 = dt1.Date
                                                                dt = DateAdd(DateInterval.Year, 10, dt)
                                                                dt1 = DateAdd(DateInterval.Year, 5, dt1)

                                                                If Not (Date.Compare(dt, dt1) > 0 Or Date.Compare(dt1, CDate("01/01/0001")) = 0) Then
                                                                    dt = dt1
                                                                End If
                                                                If Date.Compare(dt, CDate("01/01/0001")) <> 0 Then
                                                                    If Not drdsTank("TANKMODDESC").ToString.IndexOf("Cathodically Protected/Lined Interior") > -1 Then
                                                                        If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                                            If Not alTnkLinedDate.Contains(dt.ToShortDateString) Then
                                                                                If Not bolAddedSectionForOwner Then
                                                                                    AddCAPMonthlyNoticeOfTestReqHeading(headingText, drdsOwner("OWNERNAME").ToString, docTestReq, bolEnteredTextToTestReqLetter)
                                                                                    bolAddedSectionForOwner = True
                                                                                End If
                                                                                If Not bolAddedFacilityDetail Then
                                                                                    points = AddCapMontlyNoticeOfTestReqFacilityDetails(drdsFac, docTestReq, points, headingText, drdsOwner("OWNERNAME").ToString)
                                                                                    bolAddedFacilityDetail = True
                                                                                End If
                                                                                alTnkLinedDate.Add(dt.ToShortDateString)

                                                                                points = Me.AddCapMonhlyPageBreak(points + 2, points, docTestReq, headingText, drdsOwner("OWNERNAME").ToString)

                                                                                oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                                                                                'oPara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                                                                                oPara.Range.Font.Name = "Arial"
                                                                                oPara.Range.Font.Size = 10
                                                                                oPara.Range.Font.Bold = 0
                                                                                oPara.Range.Text = "Inspection of the tank interior lining must be accomplished by " + dt.ToShortDateString + "." + vbCrLf + _
                                                                                                    "Please update my records to reflect that this test was accomplished on _______________."
                                                                                oPara.Range.InsertParagraphAfter()

                                                                                InsertLines(1, docTestReq)

                                                                                points += 3
                                                                            End If
                                                                            bolEnteredTextToTestReqLetter = True
                                                                            drdsTank("MODIFIED") = True
                                                                            drdsFac("MODIFIED") = True
                                                                            drdsOwner("MODIFIED") = True
                                                                            drdsTank("LI INSPECTED") = dt
                                                                            drTank("LI INSPECTED") = dt
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If

                                                        ' IF CIU
                                                        If drdsTank("STATUS").ToString.IndexOf("Currently In Use") > -1 Then
                                                            ' TANK TT DUE / IC EXPIRES
                                                            ' Show Tank TT Due only if IC Expires is false
                                                            If Not drdsTank("TANKLD") Is DBNull.Value Then
                                                                If drdsTank("TANKLD").ToString.IndexOf("Inventory Control/Precision Tightness Testing") > -1 Then
                                                                    ' ICExpires
                                                                    dt = IIf(drdsTank("INSTALLED") Is DBNull.Value, CDate("01/01/0001"), drdsTank("INSTALLED"))
                                                                    dt1 = IIf(drdsTank("TCPINSTALLDATE") Is DBNull.Value, CDate("01/01/0001"), drdsTank("TCPINSTALLDATE"))
                                                                    dt = dt.Date
                                                                    dt1 = dt1.Date
                                                                    If Date.Compare(dt, dt1) < 0 Then
                                                                        dt = dt1
                                                                    End If
                                                                    dt1 = IIf(drdsTank("LI INSTALL") Is DBNull.Value, CDate("01/01/0001"), drdsTank("LI INSTALL"))
                                                                    dt1 = dt1.Date
                                                                    If Date.Compare(dt, dt1) < 0 Then
                                                                        dt = dt1
                                                                    End If
                                                                    dt = DateAdd(DateInterval.Year, 10, dt)

                                                                    If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                                        If Not alTnkICExpiresDate.Contains(dt.ToShortDateString) Then
                                                                            If Not bolAddedSectionForOwner Then
                                                                                AddCAPMonthlyNoticeOfTestReqHeading(headingText, drdsOwner("OWNERNAME").ToString, docTestReq, bolEnteredTextToTestReqLetter)
                                                                                bolAddedSectionForOwner = True
                                                                            End If
                                                                            If Not bolAddedFacilityDetail Then
                                                                                points = AddCapMontlyNoticeOfTestReqFacilityDetails(drdsFac, docTestReq, points, headingText, drdsOwner("OWNERNAME").ToString)
                                                                                bolAddedFacilityDetail = True
                                                                            End If
                                                                            alTnkICExpiresDate.Add(dt.ToShortDateString)

                                                                            points = Me.AddCapMonhlyPageBreak(points + 2, points, docTestReq, headingText, drdsOwner("OWNERNAME").ToString)

                                                                            oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                                                                            'oPara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                                                                            oPara.Range.Font.Name = "Arial"
                                                                            oPara.Range.Font.Size = 10
                                                                            oPara.Range.Font.Bold = 0
                                                                            oPara.Range.Text = "Please be aware that Inventory Control/Precision Tightness Testing is only a valid method of tank leak " + _
                                                                                                "detection for a period of 10 years following tank installation or upgrade. Therefore, you must choose another " + _
                                                                                                "method of tank leak detection by no later than " + dt.ToShortDateString + "."
                                                                            oPara.Range.InsertParagraphAfter()

                                                                            InsertLines(1, docTestReq)

                                                                            points += 3
                                                                        End If
                                                                        bolEnteredTextToTestReqLetter = True
                                                                        ' no need to roll over date - according to stefanie
                                                                    Else
                                                                        ' TANK TT DUE
                                                                        dt = IIf(drdsTank("TT DATE") Is DBNull.Value, CDate("01/01/0001"), drdsTank("TT DATE"))
                                                                        dt1 = IIf(drdsTank("INSTALLED") Is DBNull.Value, CDate("01/01/0001"), drdsTank("INSTALLED"))
                                                                        dt = dt.Date
                                                                        dt1 = dt1.Date

                                                                        If Date.Compare(dt, dt1) < 0 Then
                                                                            dt = dt1
                                                                        End If
                                                                        dt1 = IIf(drdsTank("TCPINSTALLDATE") Is DBNull.Value, CDate("01/01/0001"), drdsTank("TCPINSTALLDATE"))
                                                                        dt1 = dt1.Date
                                                                        If Date.Compare(dt, dt1) < 0 Then
                                                                            dt = dt1
                                                                        End If
                                                                        dt1 = IIf(drdsTank("LI INSTALL") Is DBNull.Value, CDate("01/01/0001"), drdsTank("LI INSTALL"))
                                                                        dt1 = dt1.Date
                                                                        If Date.Compare(dt, dt1) < 0 Then
                                                                            dt = dt1
                                                                        End If

                                                                        dt = DateAdd(DateInterval.Year, 5, dt)
                                                                        dt1 = CDate("12/22/1998")
                                                                        If Date.Compare(dt, dt1) < 0 Then
                                                                            dt = dt1
                                                                        End If

                                                                        If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                                            If Not alTnkTTDate.Contains(dt.ToShortDateString) Then
                                                                                If Not bolAddedSectionForOwner Then
                                                                                    AddCAPMonthlyNoticeOfTestReqHeading(headingText, drdsOwner("OWNERNAME").ToString, docTestReq, bolEnteredTextToTestReqLetter)
                                                                                    bolAddedSectionForOwner = True
                                                                                End If
                                                                                If Not bolAddedFacilityDetail Then
                                                                                    points = AddCapMontlyNoticeOfTestReqFacilityDetails(drdsFac, docTestReq, points, headingText, drdsOwner("OWNERNAME").ToString)
                                                                                    bolAddedFacilityDetail = True
                                                                                End If
                                                                                alTnkTTDate.Add(dt.ToShortDateString)

                                                                                points = Me.AddCapMonhlyPageBreak(points + 2, points, docTestReq, headingText, drdsOwner("OWNERNAME").ToString)

                                                                                oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                                                                                'oPara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                                                                                oPara.Range.Font.Name = "Arial"
                                                                                oPara.Range.Font.Size = 10
                                                                                oPara.Range.Font.Bold = 0
                                                                                oPara.Range.Text = "Precision tightness testing of the tanks must be accomplished by " + dt.ToShortDateString + "." + vbCrLf + _
                                                                                                    "Please update my records to reflect that this test was accomplished on _______________."
                                                                                oPara.Range.InsertParagraphAfter()

                                                                                InsertLines(1, docTestReq)

                                                                                points += 3
                                                                            End If
                                                                            bolEnteredTextToTestReqLetter = True
                                                                            drdsTank("MODIFIED") = True
                                                                            drdsFac("MODIFIED") = True
                                                                            drdsOwner("MODIFIED") = True
                                                                            drdsTank("TT DATE") = dt

                                                                            drTank("TT DATE") = dt
                                                                        End If
                                                                    End If

                                                                End If ' if tankld = inventory control/precision tightness testing
                                                            End If ' if tankld is null

                                                        End If ' if ciu

                                                    End If ' status is ciu / tosi
                                                End If ' status is null

                                                'Pipe Loop for CAP monthly report
                                                For Each drdsPipe In ds.Tables(3).Select("FACILITY_ID = " + drdsFac("FACILITY_ID").ToString + " AND [TANK ID] = " + drdsTank("TANK ID").ToString) ' pipe

                                                    drPipe = dtPipe.NewRow
                                                    drPipe("TANK ID") = drdsTank("TANK ID")
                                                    drPipe("PIPE ID") = drdsPipe("PIPE ID")

                                                    ' check pipe conditions
                                                    ' if pipe modified, set MODIFIED column value to true
                                                    If Not drdsPipe("STATUS") Is DBNull.Value Then
                                                        If drdsPipe("STATUS").ToString.IndexOf("Currently In Use") > -1 Or drdsPipe("STATUS").ToString.IndexOf("Temporarily Out of Service Indefinitely") > -1 Then

                                                            'DateSheerValueTest
                                                            If drdsPipe("PipeType") = 266 Then


                                                                If Not drdsPipe("DateSheerValueTest") Is DBNull.Value Then
                                                                    dt = drdsPipe("DateSheerValueTest")
                                                                    dt = dt.Date
                                                                    dt = DateAdd(DateInterval.Year, 1, dt)
                                                                    ListUnknown = dt.ToShortDateString

                                                                Else
                                                                    ListUnknown = "'Unknown'"
                                                                    dt = New Date(2009, 10, 1)
                                                                End If

                                                                If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                                    If Not alPipeSheerDate.Contains(dt.ToShortDateString) Then
                                                                        If Not bolAddedSectionForOwner Then
                                                                            AddCAPMonthlyNoticeOfTestReqHeading(headingText, drdsOwner("OWNERNAME").ToString, docTestReq, bolEnteredTextToTestReqLetter)
                                                                            bolAddedSectionForOwner = True
                                                                        End If
                                                                        If Not bolAddedFacilityDetail Then
                                                                            points = AddCapMontlyNoticeOfTestReqFacilityDetails(drdsFac, docTestReq, points, headingText, drdsOwner("OWNERNAME").ToString)
                                                                            bolAddedFacilityDetail = True
                                                                        End If
                                                                        alPipeSheerDate.Add(dt.ToShortDateString)

                                                                        points = Me.AddCapMonhlyPageBreak(points + 2, points, docTestReq, headingText, drdsOwner("OWNERNAME").ToString)

                                                                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                                                                        oPara.Range.Font.Name = "Arial"
                                                                        oPara.Range.Font.Size = 10
                                                                        oPara.Range.Font.Bold = 0
                                                                        oPara.Range.Text = "Testing of pressurized piping shear valves must be accomplished by " + ListUnknown + "." + vbCrLf + _
                                                                                            "Please update my records to reflect that this test was accomplished on _______________."
                                                                        oPara.Range.InsertParagraphAfter()

                                                                        InsertLines(1, docTestReq)

                                                                        points += 3
                                                                    End If
                                                                    bolEnteredTextToTestReqLetter = True
                                                                    drdsPipe("MODIFIED") = True
                                                                    drdsTank("MODIFIED") = True
                                                                    drdsFac("MODIFIED") = True
                                                                    drdsOwner("MODIFIED") = True
                                                                    drdsPipe("DateSheerValueTest") = dt
                                                                    drPipe("DateSheerValueTest") = dt
                                                                End If
                                                            End If


                                                            'DateSecondaryContainmentInspect
                                                            If drdsPipe("Pipe_LD_Num") = 242 Then

                                                                If Not drdsPipe("DateSecondaryContainmentInspect") Is DBNull.Value Then

                                                                    dt = drdsPipe("DateSecondaryContainmentInspect")
                                                                    dt = dt.Date
                                                                    dt = DateAdd(DateInterval.Year, 1, dt)
                                                                    ListUnknown = dt.ToShortDateString

                                                                Else
                                                                    ListUnknown = "'Unknown'"
                                                                    dt = New Date(2009, 10, 1)

                                                                End If

                                                                If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                                    If Not alPipeSecondaryDate.Contains(dt.ToShortDateString) Then
                                                                        If Not bolAddedSectionForOwner Then
                                                                            AddCAPMonthlyNoticeOfTestReqHeading(headingText, drdsOwner("OWNERNAME").ToString, docTestReq, bolEnteredTextToTestReqLetter)
                                                                            bolAddedSectionForOwner = True
                                                                        End If
                                                                        If Not bolAddedFacilityDetail Then
                                                                            points = AddCapMontlyNoticeOfTestReqFacilityDetails(drdsFac, docTestReq, points, headingText, drdsOwner("OWNERNAME").ToString)
                                                                            bolAddedFacilityDetail = True
                                                                        End If
                                                                        alPipeSecondaryDate.Add(dt.ToShortDateString)


                                                                        points = Me.AddCapMonhlyPageBreak(points + 2, points, docTestReq, headingText, drdsOwner("OWNERNAME").ToString)

                                                                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                                                                        oPara.Range.Font.Name = "Arial"
                                                                        oPara.Range.Font.Size = 10
                                                                        oPara.Range.Font.Bold = 0
                                                                        oPara.Range.Text = "Inspection of the pipe secondary containment must be accomplished by " + ListUnknown + "." + vbCrLf + _
                                                                                            "Please update my records to reflect that this test was accomplished on _______________."
                                                                        oPara.Range.InsertParagraphAfter()
                                                                        InsertLines(1, docTestReq)

                                                                        points += 3
                                                                    End If
                                                                    bolEnteredTextToTestReqLetter = True
                                                                    drdsPipe("MODIFIED") = True
                                                                    drdsTank("MODIFIED") = True
                                                                    drdsFac("MODIFIED") = True
                                                                    drdsOwner("MODIFIED") = True
                                                                    drdsPipe("DateSecondaryContainmentInspect") = dt
                                                                    drPipe("DateSecondaryContainmentInspect") = dt
                                                                End If
                                                            End If


                                                            'DateElectronicDeviceInspect
                                                            If drdsPipe("Pipe_LD_Num") = 243 Then

                                                                If Not drdsPipe("DateElectronicDeviceInspect") Is DBNull.Value Then
                                                                    dt = drdsPipe("DateElectronicDeviceInspect")
                                                                    dt = dt.Date
                                                                    dt = DateAdd(DateInterval.Year, 1, dt)
                                                                    ListUnknown = dt.ToShortDateString

                                                                Else
                                                                    ListUnknown = "'Unknown'"
                                                                    dt = New Date(2009, 10, 1)

                                                                End If

                                                                If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                                    If Not alPipeElectronicDate.Contains(dt.ToShortDateString) Then
                                                                        If Not bolAddedSectionForOwner Then
                                                                            AddCAPMonthlyNoticeOfTestReqHeading(headingText, drdsOwner("OWNERNAME").ToString, docTestReq, bolEnteredTextToTestReqLetter)
                                                                            bolAddedSectionForOwner = True
                                                                        End If
                                                                        If Not bolAddedFacilityDetail Then
                                                                            points = AddCapMontlyNoticeOfTestReqFacilityDetails(drdsFac, docTestReq, points, headingText, drdsOwner("OWNERNAME").ToString)
                                                                            bolAddedFacilityDetail = True
                                                                        End If
                                                                        alPipeElectronicDate.Add(dt.ToShortDateString)

                                                                        points = Me.AddCapMonhlyPageBreak(points + 2, points, docTestReq, headingText, drdsOwner("OWNERNAME").ToString)


                                                                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                                                                        oPara.Range.Font.Name = "Arial"
                                                                        oPara.Range.Font.Size = 10
                                                                        oPara.Range.Font.Bold = 0
                                                                        oPara.Range.Text = "Testing of the pipe electronic interstitial monitoring devices must be accomplished by " + ListUnknown + "." + vbCrLf + _
                                                                                            "Please update my records to reflect that this test was accomplished on _______________."
                                                                        oPara.Range.InsertParagraphAfter()
                                                                        InsertLines(1, docTestReq)

                                                                        points += 3
                                                                    End If
                                                                    bolEnteredTextToTestReqLetter = True
                                                                    drdsPipe("MODIFIED") = True
                                                                    drdsTank("MODIFIED") = True
                                                                    drdsFac("MODIFIED") = True
                                                                    drdsOwner("MODIFIED") = True
                                                                    drdsPipe("DateElectronicDeviceInspect") = dt
                                                                    drPipe("DateElectronicDeviceInspect") = dt
                                                                End If
                                                            End If


                                                            ' PIPE CP DATE
                                                            If Not drdsPipe("PIPE_MOD_DESC") Is DBNull.Value Then

                                                                If drdsPipe("PIPE_MOD_DESC").ToString = "Cathodically Protected" Then
                                                                    If Not drdsPipe("CP DATE") Is DBNull.Value Then

                                                                        dt = drdsPipe("CP DATE")
                                                                        dt = dt.Date
                                                                        dt = DateAdd(DateInterval.Year, 3, dt)
                                                                        ListUnknown = dt.ToShortDateString

                                                                    Else
                                                                        ListUnknown = "'Unknown'"
                                                                        dt = New Date(2009, 10, 1)

                                                                    End If

                                                                    If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                                        If Not alPipeCPDate.Contains(dt.ToShortDateString) Then
                                                                            If Not bolAddedSectionForOwner Then
                                                                                AddCAPMonthlyNoticeOfTestReqHeading(headingText, drdsOwner("OWNERNAME").ToString, docTestReq, bolEnteredTextToTestReqLetter)
                                                                                bolAddedSectionForOwner = True
                                                                            End If
                                                                            If Not bolAddedFacilityDetail Then
                                                                                points = AddCapMontlyNoticeOfTestReqFacilityDetails(drdsFac, docTestReq, points, headingText, drdsOwner("OWNERNAME").ToString)
                                                                                bolAddedFacilityDetail = True
                                                                            End If
                                                                            alPipeCPDate.Add(dt.ToShortDateString)

                                                                            points = Me.AddCapMonhlyPageBreak(points + 2, points, docTestReq, headingText, drdsOwner("OWNERNAME").ToString)

                                                                            oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                                                                            'oPara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                                                                            oPara.Range.Font.Name = "Arial"
                                                                            oPara.Range.Font.Size = 10
                                                                            oPara.Range.Font.Bold = 0
                                                                            oPara.Range.Text = "Testing of the pipe cathodic protection must be accomplished by " + ListUnknown + "." + vbCrLf + _
                                                                                                "Please update my records to reflect that this test was accomplished on _______________."
                                                                            oPara.Range.InsertParagraphAfter()

                                                                            InsertLines(1, docTestReq)

                                                                            points += 3
                                                                        End If
                                                                        bolEnteredTextToTestReqLetter = True
                                                                        drdsPipe("MODIFIED") = True
                                                                        drdsTank("MODIFIED") = True
                                                                        drdsFac("MODIFIED") = True
                                                                        drdsOwner("MODIFIED") = True
                                                                        drdsPipe("CP DATE") = dt
                                                                        drPipe("CP DATE") = dt
                                                                    End If
                                                                End If
                                                            End If

                                                            ' PIPE TERM CP DATE
                                                            Dim bolContinue As Boolean = False
                                                            If Not drdsPipe("DISP CP TYPE") Is DBNull.Value Then
                                                                If drdsPipe("DISP CP TYPE").ToString.IndexOf("Cathodically Protected") > -1 Then
                                                                    If Not drdsPipe("TERM CP TEST") Is DBNull.Value Then
                                                                        bolContinue = True
                                                                        ListUnknown = String.Empty
                                                                    Else
                                                                        bolContinue = True
                                                                        ListUnknown = "'Unknown'"
                                                                    End If
                                                                End If
                                                            End If
                                                            If Not bolContinue Then
                                                                If Not drdsPipe("TANK CP TYPE") Is DBNull.Value Then
                                                                    If drdsPipe("TANK CP TYPE").ToString.IndexOf("Cathodically Protected") > -1 Then
                                                                        If Not drdsPipe("TERM CP TEST") Is DBNull.Value Then
                                                                            bolContinue = True
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If

                                                            If bolContinue Then
                                                                If ListUnknown = String.Empty Then
                                                                    dt = drdsPipe("TERM CP TEST")
                                                                    dt = dt.Date
                                                                    dt = DateAdd(DateInterval.Year, 3, dt)
                                                                    ListUnknown = dt.ToShortDateString


                                                                Else
                                                                    ListUnknown = "'Unknown'"
                                                                    dt = New Date(2009, 10, 1)

                                                                End If
                                                                If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                                    If Not alPipeTermCPDate.Contains(dt.ToShortDateString) Then
                                                                        If Not bolAddedSectionForOwner Then
                                                                            AddCAPMonthlyNoticeOfTestReqHeading(headingText, drdsOwner("OWNERNAME").ToString, docTestReq, bolEnteredTextToTestReqLetter)
                                                                            bolAddedSectionForOwner = True
                                                                        End If
                                                                        If Not bolAddedFacilityDetail Then
                                                                            points = AddCapMontlyNoticeOfTestReqFacilityDetails(drdsFac, docTestReq, points, headingText, drdsOwner("OWNERNAME").ToString)
                                                                            bolAddedFacilityDetail = True
                                                                        End If
                                                                        alPipeTermCPDate.Add(dt.ToShortDateString)

                                                                        points = Me.AddCapMonhlyPageBreak(points + 2, points, docTestReq, headingText, drdsOwner("OWNERNAME").ToString)

                                                                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                                                                        'oPara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                                                                        oPara.Range.Font.Name = "Arial"
                                                                        oPara.Range.Font.Size = 10
                                                                        oPara.Range.Font.Bold = 0
                                                                        oPara.Range.Text = "Testing of the piping flex connector cathodic protection must be accomplished by " + ListUnknown + "." + vbCrLf + _
                                                                                            "Please update my records to reflect that this test was accomplished on _______________."
                                                                        oPara.Range.InsertParagraphAfter()

                                                                        InsertLines(1, docTestReq)

                                                                        points += 3
                                                                    End If
                                                                    bolEnteredTextToTestReqLetter = True
                                                                    drdsPipe("MODIFIED") = True
                                                                    drdsTank("MODIFIED") = True
                                                                    drdsFac("MODIFIED") = True
                                                                    drdsOwner("MODIFIED") = True
                                                                    drdsPipe("TERM CP TEST") = dt

                                                                    drPipe("TERM CP TEST") = dt
                                                                End If
                                                            End If

                                                            ' IF CIU
                                                            If drdsPipe("STATUS").ToString.IndexOf("Currently In Use") > -1 Then

                                                                ' ALLD TEST DATE
                                                                If Not drdsPipe("ALLD_TEST") Is DBNull.Value Then
                                                                    ' If drdsPipe("ALLD_TEST").ToString = "Mechanical" Then
                                                                    If drdsPipe("PIPE_TYPE_DESC").ToString = "Pressurized" And (Not (drdsPipe("PIPE_LD").ToString = "Deferred")) Then
                                                                        If Not drdsPipe("ALLD_TEST_DATE") Is DBNull.Value Then
                                                                            dt = drdsPipe("ALLD_TEST_DATE")
                                                                            dt = dt.Date
                                                                            dt = DateAdd(DateInterval.Year, 1, dt)
                                                                            ListUnknown = dt.ToShortDateString

                                                                        Else
                                                                            ListUnknown = "'Unknown'"
                                                                            dt = New Date(2009, 10, 1)


                                                                        End If

                                                                        If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                                            If Not alPipeALLDDate.Contains(dt.ToShortDateString) Then
                                                                                If Not bolAddedSectionForOwner Then
                                                                                    AddCAPMonthlyNoticeOfTestReqHeading(headingText, drdsOwner("OWNERNAME").ToString, docTestReq, bolEnteredTextToTestReqLetter)
                                                                                    bolAddedSectionForOwner = True
                                                                                End If
                                                                                If Not bolAddedFacilityDetail Then
                                                                                    points = AddCapMontlyNoticeOfTestReqFacilityDetails(drdsFac, docTestReq, points, headingText, drdsOwner("OWNERNAME").ToString)
                                                                                    bolAddedFacilityDetail = True
                                                                                End If
                                                                                alPipeALLDDate.Add(dt.ToShortDateString)

                                                                                points = Me.AddCapMonhlyPageBreak(points + 2, points, docTestReq, headingText, drdsOwner("OWNERNAME").ToString)

                                                                                oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                                                                                'oPara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                                                                                oPara.Range.Font.Name = "Arial"
                                                                                oPara.Range.Font.Size = 10
                                                                                oPara.Range.Font.Bold = 0
                                                                                oPara.Range.Text = "Testing of the automatic line leak detector must be accomplished by " + ListUnknown + "." + vbCrLf + _
                                                                                                    "Please update my records to reflect that this test was accomplished on _______________."
                                                                                oPara.Range.InsertParagraphAfter()

                                                                                InsertLines(1, docTestReq)

                                                                                points += 3
                                                                            End If
                                                                            bolEnteredTextToTestReqLetter = True
                                                                            drdsPipe("MODIFIED") = True
                                                                            drdsTank("MODIFIED") = True
                                                                            drdsFac("MODIFIED") = True
                                                                            drdsOwner("MODIFIED") = True
                                                                            drdsPipe("ALLD_TEST_DATE") = dt

                                                                            drPipe("ALLD_TEST_DATE") = dt
                                                                        End If
                                                                    End If
                                                                End If

                                                                ' PIPE LINE
                                                                If Not drdsPipe("PIPE_LD") Is DBNull.Value Then
                                                                    If drdsPipe("PIPE_LD").ToString = "Line Tightness Testing" Then
                                                                        If Not drdsPipe("PIPE_TYPE_DESC") Is DBNull.Value Then

                                                                            ' US SUCTION
                                                                            If drdsPipe("PIPE_TYPE_DESC").ToString = "U.S. Suction" Then
                                                                                If Not drdsPipe("TT DATE") Is DBNull.Value Then

                                                                                    dt = drdsPipe("TT DATE")
                                                                                    dt = dt.Date
                                                                                    dt = DateAdd(DateInterval.Year, 3, dt)
                                                                                    ListUnknown = dt.ToShortDateString

                                                                                Else
                                                                                    ListUnknown = "'Unknown'"
                                                                                    dt = New Date(2009, 10, 1)

                                                                                End If


                                                                                If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                                                    If Not alPipeLineUSDate.Contains(dt.ToShortDateString) Then
                                                                                        If Not bolAddedSectionForOwner Then
                                                                                            AddCAPMonthlyNoticeOfTestReqHeading(headingText, drdsOwner("OWNERNAME").ToString, docTestReq, bolEnteredTextToTestReqLetter)
                                                                                            bolAddedSectionForOwner = True
                                                                                        End If
                                                                                        If Not bolAddedFacilityDetail Then
                                                                                            points = AddCapMontlyNoticeOfTestReqFacilityDetails(drdsFac, docTestReq, points, headingText, drdsOwner("OWNERNAME").ToString)
                                                                                            bolAddedFacilityDetail = True
                                                                                        End If
                                                                                        alPipeLineUSDate.Add(dt.ToShortDateString)

                                                                                        points = Me.AddCapMonhlyPageBreak(points + 2, points, docTestReq, headingText, drdsOwner("OWNERNAME").ToString)

                                                                                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                                                                                        'oPara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                                                                                        oPara.Range.Font.Name = "Arial"
                                                                                        oPara.Range.Font.Size = 10
                                                                                        oPara.Range.Font.Bold = 0
                                                                                        oPara.Range.Text = "Precision tightness testing of the 'U.S.' suction piping must be accomplished by " + ListUnknown + "." + vbCrLf + _
                                                                                                            "Please update my records to reflect that this test was accomplished on _______________."
                                                                                        oPara.Range.InsertParagraphAfter()

                                                                                        InsertLines(1, docTestReq)
                                                                                        points += 3
                                                                                    End If
                                                                                    bolEnteredTextToTestReqLetter = True
                                                                                    drdsPipe("MODIFIED") = True
                                                                                    drdsTank("MODIFIED") = True
                                                                                    drdsFac("MODIFIED") = True
                                                                                    drdsOwner("MODIFIED") = True
                                                                                    drdsPipe("TT DATE") = dt

                                                                                    drPipe("TT DATE") = dt
                                                                                End If

                                                                                ' PRESSURIZED
                                                                            ElseIf drdsPipe("PIPE_TYPE_DESC").ToString = "Pressurized" Then
                                                                                If Not drdsPipe("TT DATE") Is DBNull.Value Then
                                                                                    dt = drdsPipe("TT DATE")
                                                                                    dt = dt.Date
                                                                                    dt = DateAdd(DateInterval.Year, 1, dt)
                                                                                    ListUnknown = dt.ToShortDateString

                                                                                Else
                                                                                    ListUnknown = "'Unknown'"
                                                                                    dt = New Date(2009, 10, 1)


                                                                                End If

                                                                                If Date.Compare(dt, dtProcessingStart) >= 0 And Date.Compare(dtProcessingEnd, dt) >= 0 Then
                                                                                    If Not alPipeLinePressDate.Contains(dt.ToShortDateString) Then
                                                                                        If Not bolAddedSectionForOwner Then
                                                                                            AddCAPMonthlyNoticeOfTestReqHeading(headingText, drdsOwner("OWNERNAME").ToString, docTestReq, bolEnteredTextToTestReqLetter)
                                                                                            bolAddedSectionForOwner = True
                                                                                        End If
                                                                                        If Not bolAddedFacilityDetail Then
                                                                                            points = AddCapMontlyNoticeOfTestReqFacilityDetails(drdsFac, docTestReq, points, headingText, drdsOwner("OWNERNAME").ToString)
                                                                                            bolAddedFacilityDetail = True
                                                                                        End If
                                                                                        alPipeLinePressDate.Add(dt.ToShortDateString)

                                                                                        points = Me.AddCapMonhlyPageBreak(points + 2, points, docTestReq, headingText, drdsOwner("OWNERNAME").ToString)

                                                                                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                                                                                        'oPara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                                                                                        oPara.Range.Font.Name = "Arial"
                                                                                        oPara.Range.Font.Size = 10
                                                                                        oPara.Range.Font.Bold = 0
                                                                                        oPara.Range.Text = "Precision tightness testing of the pressurized piping must be accomplished by " + ListUnknown + "." + vbCrLf + _
                                                                                                            "Please update my records to reflect that this test was accomplished on _______________."
                                                                                        oPara.Range.InsertParagraphAfter()

                                                                                        InsertLines(1, docTestReq)

                                                                                        points += 3
                                                                                    End If
                                                                                    bolEnteredTextToTestReqLetter = True
                                                                                    drdsPipe("MODIFIED") = True
                                                                                    drdsTank("MODIFIED") = True
                                                                                    drdsFac("MODIFIED") = True
                                                                                    drdsOwner("MODIFIED") = True
                                                                                    drdsPipe("TT DATE") = dt

                                                                                    drPipe("TT DATE") = dt
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If

                                                            End If ' if ciu

                                                        End If ' if status is ciu / tosi
                                                    End If ' status is null

                                                    If drdsPipe("MODIFIED") = True Then
                                                        dtPipe.Rows.Add(drPipe)
                                                    End If

                                                Next ' pipe

                                                If drdsTank("MODIFIED") = True Then
                                                    dtTank.Rows.Add(drTank)
                                                End If

                                            Next ' tank

                                            ' Save Assistance doc by tank
                                            Dim keepSaving As Boolean = True
                                            Dim cnt As Integer = 0

                                            While keepSaving
                                                Try

                                                    ' If i Mod 30 = 0 Or i = ds.Tables(0).Rows.Count - 1 Then
                                                    ' .Save()
                                                    ' End If

                                                    keepSaving = False

                                                Catch ex As Exception

                                                    cnt += 1


                                                    If ex.ToString.ToUpper.IndexOf(" PERMISSION") > -1 Then

                                                        If cnt >= 10 Then
                                                            Threading.Thread.Sleep(2000)
                                                            cnt = 0

                                                        End If

                                                        keepSaving = True

                                                    End If

                                                End Try

                                            End While

                                        End With  ' With docTestReq

                                        If drdsFac("MODIFIED") = True Then
                                            dtFac.Rows.Add(drFac)
                                        End If

                                    Next ' facility

                                    ' create assistance letter for owner
                                    If bolAddedSectionForOwner Then
                                        AddCAPMonthlyComplianceAssistanceLetter(drdsOwner, docAssist, bolEnteredTextToAssistLetter, strAssistTemplate)
                                        bolEnteredTextToAssistLetter = True
                                        'Add Owner IDs for printing labels
                                        If strOwnerIDs = String.Empty Then
                                            strOwnerIDs = drdsOwner("OWNER_ID").ToString.Trim
                                        Else
                                            strOwnerIDs += "," + drdsOwner("OWNER_ID").ToString.Trim
                                        End If

                                    End If

                                    ' Save Assistance doc
                                    Dim keepSaving2 As Boolean = True
                                    Dim cnt2 As Integer = 0

                                    While keepSaving2
                                        Try

                                            'If i Mod 30 = 0 Or i = ds.Tables(0).Rows.Count - 1 Then
                                            '.Save()
                                            'End If

                                            keepSaving2 = False

                                        Catch ex As Exception

                                            cnt2 += 1


                                            If ex.ToString.ToUpper.IndexOf(" PERMISSION") > -1 Then

                                                If cnt2 >= 10 Then
                                                    Threading.Thread.Sleep(2000)
                                                    cnt2 = 0


                                                End If

                                                keepSaving2 = True

                                            End If

                                        End Try

                                    End While


                                End With ' With docAssist

                                If drdsOwner("MODIFIED") = True Then
                                    dtOwner.Rows.Add(drOwner)
                                End If
                            End If


                        Next ' owner



                        'dtEnd = DateTime.Now
                        'ts = dtEnd.Subtract(dtStart)
                        'strTime += vbCrLf + "generate cap monthly report: " + ts.ToString

                        If bolEnteredTextToAssistLetter Or bolEnteredTextToTestReqLetter Then
                            docTestReq.Activate()

                            ''' roll over dates
                            'If dtOwner.Rows.Count > 0 Then

                            'Dim dtCPDate, dtLIInspected, dtTTDate, dtTermCpTest, dtAlldTestDate, dtSpillTested, dtOverfillInspected, dtTankSecondary, dtTankElectronic, dtATG, dtShear, dtPipeSecondary, dtPipeElectronic As Date

                            ''''dtStart = DateTime.Now


                            'For Each drOwner In dtOwner.Rows ' owner
                            '  For Each drFac In dtFac.Select("OWNER_ID = " + drOwner("OWNER_ID").ToString) ' facility
                            ''' tank
                            '  For Each drTank In dtTank.Select("FACILITY_ID = " + drFac("FACILITY_ID").ToString)
                            '''  ' if all the dates are null, this row is a place holder, no need to roll over dates
                            '  If Not (drTank("CP DATE") Is DBNull.Value And drTank("LI INSPECTED") Is DBNull.Value And drTank("TT DATE") Is DBNull.Value And drTank("DateSpillPreventionLastTested") Is DBNull.Value And drTank("DateOverfillPreventionLastInspected") Is DBNull.Value And drTank("DateSecondaryContainmentLastInspected") Is DBNull.Value And drTank("DateElectronicDeviceInspected") Is DBNull.Value And drTank("DateATGLastInspected") Is DBNull.Value) Then
                            ' dtCPDate = IIf(drTank("CP DATE") Is DBNull.Value, CDate("01/01/0001"), drTank("CP DATE"))
                            '   dtLIInspected = IIf(drTank("LI INSPECTED") Is DBNull.Value, CDate("01/01/0001"), drTank("LI INSPECTED"))
                            '   dtTTDate = IIf(drTank("TT DATE") Is DBNull.Value, CDate("01/01/0001"), drTank("TT DATE"))
                            '   dtSpillTested = IIf(drTank("DateSpillPreventionLastTested") Is DBNull.Value, CDate("01/01/0001"), drTank("DateSpillPreventionLastTested"))
                            '   dtOverfillInspected = IIf(drTank("DateOverfillPreventionLastInspected") Is DBNull.Value, CDate("01/01/0001"), drTank("DateOverfillPreventionLastInspected"))
                            '   dtTankSecondary = IIf(drTank("DateSecondaryContainmentLastInspected") Is DBNull.Value, CDate("01/01/0001"), drTank("DateSecondaryContainmentLastInspected"))
                            '   dtTankElectronic = IIf(drTank("DateElectronicDeviceInspected") Is DBNull.Value, CDate("01/01/0001"), drTank("DateElectronicDeviceInspected"))
                            '   dtATG = IIf(drTank("DateATGLastInspected") Is DBNull.Value, CDate("01/01/0001"), drTank("DateATGLastInspected"))

                            '   pOwn.RollOverTankCAPDates(drFac("FACILITY_ID"), drTank("TANK ID"), dtCPDate, dtLIInspected, dtTTDate, dtSpillTested, dtOverfillInspected, dtTankSecondary, dtTankElectronic, dtATG, MusterContainer.AppUser.ID)
                            ' End If

                            ''' pipe
                            '  For Each drPipe In dtPipe.Select("[TANK ID] = " + drTank("TANK ID").ToString) ' pipe
                            '     dtCPDate = IIf(drPipe("CP DATE") Is DBNull.Value, CDate("01/01/0001"), drPipe("CP DATE"))
                            '     dtTermCpTest = IIf(drPipe("TERM CP TEST") Is DBNull.Value, CDate("01/01/0001"), drPipe("TERM CP TEST"))
                            '    dtAlldTestDate = IIf(drPipe("ALLD_TEST_DATE") Is DBNull.Value, CDate("01/01/0001"), drPipe("ALLD_TEST_DATE"))
                            '   dtTTDate = IIf(drPipe("TT DATE") Is DBNull.Value, CDate("01/01/0001"), drPipe("TT DATE"))
                            '  dtShear = IIf(drPipe("DateSheerValueTest") Is DBNull.Value, CDate("01/01/0001"), drPipe("DateSheerValueTest"))
                            '  dtPipeSecondary = IIf(drPipe("DateSecondaryContainmentInspect") Is DBNull.Value, CDate("01/01/0001"), drPipe("DateSecondaryContainmentInspect"))
                            ''  dtPipeElectronic = IIf(drPipe("DateElectronicDeviceInspect") Is DBNull.Value, CDate("01/01/0001"), drPipe("DateElectronicDeviceInspect"))
                            ' pOwn.RollOverPipeCAPDates(drFac("FACILITY_ID"), drPipe("PIPE ID"), dtCPDate, dtTermCpTest, dtAlldTestDate, dtTTDate, dtShear, dtPipeSecondary, dtPipeElectronic, MusterContainer.AppUser.ID)
                            '  Next ' pipe

                            '     Next ' tank
                            '    Next ' facility
                            '   Next ' owner

                            'End If


                            .Visible = True
                            ' Save to print basket
                            UIUtilsGen.SaveDocument(0, 0, strAssistDocName, "CAP Monthly Assistance Letter", doc_path, "CAP Monthly Assistance Letter for " + processingMonthYear.Month.ToString + "-" + processingMonthYear.Year.ToString, UIUtilsGen.ModuleID.CAPProcess, 0, 0, 0)
                            UIUtilsGen.SaveDocument(0, 0, strTestReqDocName, "CAP Monthly Processing Report", doc_path, "CAP Monthly Processing Report for " + processingMonthYear.Month.ToString + "-" + processingMonthYear.Year.ToString, UIUtilsGen.ModuleID.CAPProcess, 0, 0, 0)

                            'Insert all the Owner details for printing Labels
                            pOwn.RunSQLQuery("EXEC spPutCAPLabels '" + strOwnerIDs + "','" + processingMonthYear + "','Monthly'")


                        Else
                            MsgBox("No Records Found")
                            ' delete the docs created at the top as no text was entered to the files
                            bolDeleteFilesCreated = True
                            Exit Sub
                        End If

                    End With ' with wordapp
                End If

            Else
                MsgBox("No Records Found")
                ' delete the docs created at the top as no text was entered to the files
                bolDeleteFilesCreated = True
                Exit Sub
            End If

        Catch ex As Exception
            bolDeleteFilesCreated = True
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            'dtEndAll = DateTime.Now
            'ts = dtEndAll.Subtract(dtStartAll)
            'strTime += vbCrLf + "complete process: " + ts.ToString
            'MsgBox(strTime)
            If bolDeleteFilesCreated Then
                If Not WordApp Is Nothing Then
                    If Not docAssist Is Nothing Then
                        docAssist.Close(False)
                        UIUtilsGen.Delay(, 2)
                        System.IO.File.Delete(doc_path + strAssistDocName)
                    End If
                    If Not docTestReq Is Nothing Then
                        docTestReq.Close(False)
                        UIUtilsGen.Delay(, 2)
                        System.IO.File.Delete(doc_path + strTestReqDocName)
                    End If
                End If
            End If
        End Try
    End Sub

End Class
