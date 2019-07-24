Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared



'Fixes and upgrades
'     1.1             Thomas Franey          2/20/2009             Added the ability to nullify begin and end dates
''                                                                 line: 264 - 279 used
'                      Hua Cao               06/13/2011            Change getting UserID to User Name; Add "Phone" field



Public Class CrystalReportsParamForm
    Inherits System.Windows.Forms.Form

#Region "Private & Public members"
    Dim strError As String
    Public Cr As ReportDocument
    Public Report As MUSTER.Info.ReportInfo
    Public ObjList() As Object = Nothing
    Private subReportIndex As Integer = -1
    Private ReportParams As MUSTER.BusinessLogic.pReportParams
#End Region


#Region "Events"
    Friend Event ReturnReport()
#End Region


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    Sub New(ByVal objs() As Object)

        ObjList = objs

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents btnDone As System.Windows.Forms.Button
    Friend WithEvents lblMsg As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnDone = New System.Windows.Forms.Button
        Me.lblMsg = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'btnDone
        '
        Me.btnDone.Location = New System.Drawing.Point(112, 72)
        Me.btnDone.Name = "btnDone"
        Me.btnDone.Size = New System.Drawing.Size(176, 24)
        Me.btnDone.TabIndex = 0
        Me.btnDone.Text = "Generate Report"
        '
        'lblMsg
        '
        Me.lblMsg.Location = New System.Drawing.Point(32, 8)
        Me.lblMsg.Name = "lblMsg"
        Me.lblMsg.Size = New System.Drawing.Size(344, 16)
        Me.lblMsg.TabIndex = 1
        Me.lblMsg.Text = "You must fill out all form items below before generating the report."
        '
        'CrystalReportsParamForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(416, 102)
        Me.Controls.Add(Me.lblMsg)
        Me.Controls.Add(Me.btnDone)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "CrystalReportsParamForm"
        Me.Text = "Report Parameters"
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Constants & Enumerators"
    Enum pType

        pString = 0
        pInt = 1
        PMoney = 2
        pDate = 3
        pLogical = 4
        pErr = 5

    End Enum

#End Region


#Region "Private Methods"
    Private Function getParameterType(ByVal param As ParameterFieldDefinition) As pType

        Dim thisPType As pType = pType.pString

        With param.ParameterValueKind

            If InStr(.ToString, "Number") Then
                thisPType = pType.pInt
            ElseIf InStr(.ToString, "String") Then
                thisPType = pType.pString
            ElseIf InStr(.ToString, "DateP") OrElse InStr(.ToString, "DateTime") Then
                thisPType = pType.pDate
            ElseIf InStr(.ToString, "Boolean") Then
                thisPType = pType.pLogical
            ElseIf InStr(.ToString, "Currency") Then
                thisPType = pType.PMoney
            Else
                thisPType = pType.pErr
            End If
        End With

        Return thisPType
    End Function

    Private Sub CrystalReportsTestParamForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim intTop As Integer = 0
        Dim Reports As New MUSTER.BusinessLogic.pReport

        ReportParams = New MUSTER.BusinessLogic.pReportParams

        ReportParams.ReportID = Report.ID 'Cr.FileName.Substring(Cr.FileName.LastIndexOf("\") + 1, Cr.FileName.Length - (Cr.FileName.LastIndexOf("\") + 1))

        For i = 0 To Cr.DataDefinition.ParameterFields.Count - 1

            Dim param As ParameterFieldDefinition = Cr.DataDefinition.ParameterFields(i)
            Dim paramField As ParameterField = Cr.ParameterFields(i)
            Dim label As New Label

            If param.Name.Length = 0 OrElse param.Name.IndexOf("UserID") = -1 OrElse param.Name.IndexOf("UserID") > 1 Then
                If Not (param.Name.Length = 0 OrElse param.Name.IndexOf("Phone") = -1 OrElse param.Name.IndexOf("Phone") > 1) Then
                    Dim crParameterValues As ParameterValues

                    crParameterValues = param.CurrentValues

                    crParameterValues.AddValue(MusterContainer.AppUser.PhoneNumber)

                    param.ApplyCurrentValues(crParameterValues)
                Else
                    ReportParams.Retrieve(String.Format("SYSTEM|REPORTPARAMS|{0}|{1} ", ReportParams.ReportID, param.Name), False)

                    With label
                        If paramField.PromptText <> "" Then
                            .Text = paramField.PromptText
                        Else
                            .Text = String.Format("{0} {1}", param.Name(), " (undefined) ")
                        End If

                        .Left = 30
                        .Top = intTop + 50
                        .Height = 50
                        .Width = 150

                    End With

                    'Validation to skip sub-report parameters.
                    'If (Cr.DataDefinition.ParameterFields(i).ReportName = "rptTec_LustStatusReport_Activity") Or _
                    '        (Cr.DataDefinition.ParameterFields(i).ReportName = "rptTec_LustStatusReport_Document") Or _
                    '        (Cr.DataDefinition.ParameterFields(i).ReportName = "rptTec_LustStatusReport_Comments") Then
                    '    Exit For
                    'End If

                    If Not param.ReportName = String.Empty Then
                        subReportIndex = i
                        Exit For
                    End If

                    Controls.Add(label)

                    Dim strName As String
                    Dim strType As String
                    Dim pTypeVal As pType = getParameterType(param)

                    Select Case pTypeVal
                        Case pType.pInt
                            strName = "int_"
                        Case pType.pString
                            strName = "str_"
                        Case pType.pDate
                            strName = "dat_"
                        Case pType.pLogical
                            strName = "boo_"
                        Case pType.PMoney
                            strName = "cur_"
                        Case Else
                            strName = "err_"
                    End Select

                    Select Case pTypeVal
                        Case pType.pDate
                            strType = "dat"
                        Case pType.pLogical
                            strType = "boo"
                        Case Else
                            strType = "str"
                    End Select

                    strName = String.Format("{0}{1}", strName, param.Name)


                    'Get the default parameters.
                    Dim paramValues As ParameterDiscreteValue
                    Dim paramValueDesc As String = String.Empty
                    Dim count As Integer = 0
                    Dim dtParamTable As Data.DataTable
                    Dim row As Data.DataRow

                    'TODO 1 - check for combobox values
                    'TODO 2- check for multiple default values
                    'TODO 3- check for single default values

                    If param.DefaultValues.Count > 1 Then

                        dtParamTable = CreateTableCombo()

                        For count = 0 To param.DefaultValues.Count - 1

                            paramValues = param.DefaultValues.Item(count)
                            paramValueDesc = paramValues.Description

                            row = dtParamTable.NewRow

                            If Not paramValues.Description Is Nothing Then
                                row.Item("PropertyID") = IIf(paramValueDesc.IndexOf(".rpt") > -1, paramValueDesc, paramValues.Value)
                                row("PropertyName") = IIf(paramValueDesc = String.Empty OrElse paramValueDesc.IndexOf(".rpt") > -1, paramValues.Value, paramValueDesc)
                            Else

                                row.Item("PropertyID") = paramValues.Value
                                row("PropertyName") = paramValues.Value

                            End If

                            dtParamTable.Rows.Add(row)
                        Next


                        Dim cmbStaffID As New ComboBox
                        With cmbStaffID

                            .Name = strName

                            .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
                            .Left = 180
                            .Top = intTop + 50
                            .Width = 160
                            .DataSource = dtParamTable.DefaultView
                            .DisplayMember = "PropertyName"
                            .ValueMember = "PropertyID"
                        End With

                        AddHandler cmbStaffID.SelectedIndexChanged, AddressOf SelectionChange
                        Controls.Add(cmbStaffID)

                    ElseIf strType = "str" Then

                        'validation to add combo box.
                        ' If (Report.FileName = "InspectorFacilities.rpt") Or _
                        '(Report.FileName = "OutstandingCheckList.rpt") Or _
                        '(Report.FileName = "ScheduledInspectionsByInspector.rpt") Then

                        Dim ComboParam As String = Cr.DataDefinition.ParameterFields(i).Name.ToString

                        If ComboParam.Substring(ComboParam.Length - 5, 5) = "COMBO" Then

                            Dim strQuery As String = String.Empty
                            Dim defaultValue As ParameterDiscreteValue = param.DefaultValues.Item(0)

                            If getParameterType(param) = pType.pString Then

                                If defaultValue.Value = "0" Or defaultValue.Value = String.Empty Or _
                                   defaultvalue.Value = " " Or defaultvalue.Value = "- ALL -" Then

                                    strQuery = defaultValue.Description
                                Else
                                    strQuery = defaultValue.Value
                                End If

                            ElseIf getParameterType(param) = pType.pInt Then
                                strQuery = defaultValue.Description
                            End If

                            Dim cmbStaffIDs As New ComboBox

                            With cmbStaffIDs
                                .Name = strName
                                .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
                                .Sorted = False
                                .Left = 180
                                .Top = intTop + 50
                                .Width = 160
                                .Tag = strQuery
                            End With

                            UpdateComboQuery(cmbStaffIDs)
                            Controls.Add(cmbStaffIDs)

                        Else

                            Dim textbox As New TextBox

                            With textbox
                                .Name = strName
                                .Left = 190
                                .Top = intTop + 50
                            End With


                            'Set the Default Value for string and integer.
                            If param.DefaultValues.Count > 0 Then

                                Dim defaultValue As ParameterDiscreteValue = param.DefaultValues.Item(0)

                                If getParameterType(param) = pType.pString Or getParameterType(param) = pType.pInt Then
                                    textbox.Text = defaultValue.Value
                                End If

                            End If

                            Controls.Add(textbox)

                        End If

                    ElseIf strType = "boo" Then

                        Dim rdlist As New RadioButton

                        With rdlist

                            .Name = strName
                            .Text = "TRUE"
                            .Checked = True
                            .Left = 190
                            .Top = intTop + 50

                            Controls.Add(rdlist)
                        End With


                        rdlist = New RadioButton

                        With rdlist
                            .Name = strName
                            .Text = "FALSE"
                            .Checked = False
                            .Left = 190
                            .Top = intTop + 70

                            Controls.Add(rdlist)

                        End With

                    ElseIf strType = "dat" Then

                        Dim top As Integer = intTop + 50
                        Dim dtPicker As New DateTimePicker

                        If strName.ToUpper.IndexOf("BEGIN") > -1 OrElse strName.ToUpper.IndexOf("END") >= -1 Then

                            Dim dtCBox As New CheckBox

                            With dtCBox
                                .Name = String.Format("{0}ChkBox", strName)
                                .Left = 190
                                .Top = top
                                .Text = "All DateRange"
                                .Tag = dtPicker
                            End With

                            top = top + 40

                            Controls.Add(dtCBox)
                            AddHandler dtCBox.CheckedChanged, AddressOf DisableCheckBox

                        End If
                        dtPicker.Name = strName
                        dtPicker.Left = 190
                        dtPicker.Top = top

                        intTop += 40

                        'Set the Default Value for string and integer.
                        If param.DefaultValues.Count > 0 Then

                            Dim defaultValue As ParameterDiscreteValue = param.DefaultValues.Item(0)

                            If getParameterType(param) = pType.pDate Then
                                dtPicker.Value = defaultvalue.Value
                            End If

                        End If

                        Controls.Add(dtPicker)

                    Else

                        Dim textbox As New TextBox

                        With textbox
                            .Name = strName
                            .Left = 190
                            .Top = intTop + 50
                            Controls.Add(textbox)
                        End With

                    End If

                    Cr.DataDefinition.ParameterFields.MoveNext()

                    intTop = intTop + 50
                End If
            Else

                    Dim crParameterValues As ParameterValues

                    crParameterValues = param.CurrentValues

                    crParameterValues.AddValue(MusterContainer.AppUser.Name)

                    param.ApplyCurrentValues(crParameterValues)

                End If

        Next



        'Set the Focus to the first parameter control.
        Dim item As Control

        For Each item In MyBase.Controls

            If InStr(item.GetType.ToString, "TextBox") Or _
                    InStr(LCase(item.GetType.ToString), "radiobutton") Or _
                    InStr(LCase(item.GetType.ToString), "datetime") Or _
                    InStr(LCase(item.GetType.ToString), "combobox") Then

                item.Select()
                Exit For
            End If

        Next

        btnDone.Top = intTop + 50 '100

        If intTop + 125 > 700 Then
            Height = 700
        Else
            Height = intTop + 125 '200
        End If

    End Sub

    Private Sub UpdateComboQuery(ByVal cmbStaffIDs As ComboBox)

        Dim strQuery As String = cmbStaffIDs.Tag
        Dim dtTableCombo As DataTable = CreateTableCombo()
        Dim value As Object = Nothing
        Dim ValueID As Object = Nothing

        While strQuery.IndexOf("@") > -1

            Dim param As String = strQuery.Substring(strQuery.IndexOf("@"))
            Dim stopIndex As Integer = IIf(param.IndexOf(" ") = -1, IIf(param.IndexOf(")") = -1, -1, param.IndexOf(")")), param.IndexOf(" "))

            If stopIndex > 0 Then
                param = param.Substring(0, stopIndex - 1)
            End If

            For Each cntrl As Control In Me.Controls

                If cntrl.Name.Substring(cntrl.Name.IndexOf("_") + 1).Replace("@", String.Empty) = param.Substring(1).Trim Then

                    If TypeOf cntrl Is ComboBox Then
                        value = DirectCast(cntrl, ComboBox).Text
                        ValueID = DirectCast(cntrl, ComboBox).SelectedValue

                    ElseIf TypeOf cntrl Is TextBox Then
                        value = DirectCast(cntrl, TextBox).Text
                        ValueID = DirectCast(cntrl, TextBox).Text

                    End If

                    Exit For
                End If

            Next

            If value Is Nothing OrElse TypeOf value Is DBNull Then

                strQuery = strQuery.Replace(param, "null")
            Else
                strQuery = strQuery.Replace(param, value)

            End If

        End While


        dtTableCombo = MusterContainer.pOwn.RunSQLQuery(strQuery).Tables(0)
        Dim drow As DataRow

        drow = dtTableCombo.NewRow
        drow("PropertyName") = "- ALL -"
        'dtTableCombo.DefaultView.Sor = "PropertyName"
        cmbStaffIDs.DisplayMember = "PropertyName"
        cmbStaffIDs.DataSource = dtTableCombo.DefaultView



        If dtTableCombo.Columns.Count > 1 Then
            drow("PropertyID") = 0
            cmbStaffIDs.ValueMember = "PropertyID"
        Else
            cmbStaffIDs.ValueMember = "PropertyName"
        End If

        dtTableCombo.Rows.InsertAt(drow, 0)


    End Sub

    Private Function CreateTableCombo() As Data.DataTable
        Try

            Dim dtParamTable As New Data.DataTable
            Dim row As Data.DataRow
            Dim col1 As New Data.DataColumn
            Dim col2 As New Data.DataColumn

            col1.ColumnName = "PropertyID"
            col2.ColumnName = "PropertyName"

            dtParamTable.Columns.Add(col1)
            dtParamTable.Columns.Add(col2)

            Return dtParamTable

        Catch ex As Exception

            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Function

    Public Function CheckValue(ByVal name As String, ByVal value As String, ByVal offset As Integer) As Boolean

        Dim fieldname As String
        Dim label As New Label

        ReportParams.Retrieve(String.Format("SYSTEM|REPORTPARAMS|{0}|{1}", ReportParams.ReportID, Microsoft.VisualBasic.Right(name, Len(name) - offset)), False)

        If ReportParams.ParamDescription <> String.Empty Then
            fieldname = ReportParams.ParamDescription
        Else
            fieldname = String.Format("{0} (undefined)", Microsoft.VisualBasic.Right(name, Len(name) - offset))
        End If

        strError = String.Format("{0}{1}{2}", strError, vbCrLf, fieldname)

        If value <> String.Empty AndAlso offset > 0 Then


            Select Case Microsoft.VisualBasic.Left(name, offset - 1)

                Case "str"
                    If InStr(value, "'") Then

                        strError = String.Format("{0} was an invalid string.", strError)
                        Return False
                    Else
                        strError = String.Empty
                        Return True
                    End If

                Case "int"
                    If Not IsNumeric(value) Then

                        strError = String.Format("{0} was an invalid integer.", strError)
                        Return False
                    Else
                        strError = String.Empty
                        Return True
                    End If

                Case "dat"
                    If Not IsDate(value) Then

                        strError = String.Format("{0} was an invalid date.", strError)
                        Return False
                    Else
                        strError = String.Empty
                        Return True
                    End If

                Case Else
                    strError = String.Empty
                    Return True
            End Select

        ElseIf value <> String.Empty Then
            strError = String.Empty
            Return True

        End If

    End Function


#End Region


#Region "Control Events"

    Private Sub SelectionChange(ByVal sender As Object, ByVal e As EventArgs)

        For Each cntrl As Control In Me.Controls

            If TypeOf cntrl Is ComboBox AndAlso cntrl.Tag <> Nothing Then
                UpdateComboQuery(cntrl)
            End If

        Next

    End Sub



    Public Sub PushParamters()

        strError = String.Empty
        Dim item As Control
        Dim booValid As Boolean = True


        Dim cnt As Integer = 0


        ReportParams = New MUSTER.BusinessLogic.pReportParams

        ReportParams.ReportID = Report.ID 'Cr.FileName.Substr

        If booValid = True Then

            For Each obj As Object In ObjList

                If obj Is Nothing Then obj = "0"

                If obj Is DBNull.Value Then obj = "0"

                Dim strValue As String = obj.ToString

                If CheckValue(Cr.DataDefinition.ParameterFields.Item(cnt).Name, strValue, 0) = True Then
                    Dim strName As String = Microsoft.VisualBasic.Right(Cr.DataDefinition.ParameterFields.Item(cnt).Name, _
                                            Len(Cr.DataDefinition.ParameterFields.Item(cnt).Name) - 0)

                    '
                    'RaiseEvent ReturnParameters(colParams)
                    ' Declare the parameter related objects.
                    '

                    Dim crParameterDiscreteValue As ParameterDiscreteValue
                    Dim crParameterFieldDefinitions As ParameterFieldDefinitions
                    Dim crParameterFieldLocation As ParameterFieldDefinition
                    Dim crParameterValues As ParameterValues
                    '
                    ' Get the report's parameters collection.

                    crParameterFieldDefinitions = Cr.DataDefinition.ParameterFields
                    '
                    ' Set the first parameter
                    ' - Get the parameter, tell it to use the current values vs default value.
                    ' - Tell it the parameter contains 1 discrete value vs multiple values.
                    ' - Set the parameter's value.
                    ' - Add it and apply it.
                    ' - Repeat these statements for each parameter.
                    '
                    crParameterValues = crParameterFieldDefinitions.Item(strName).CurrentValues

                    crParameterValues.AddValue(strValue)


                    Cr.DataDefinition.ParameterFields.Item(strName).ApplyCurrentValues(crParameterValues)


                End If
                cnt += 1

            Next

            RaiseEvent ReturnReport()

        Else

            MsgBox(strError)
        End If



    End Sub

    Private Sub btnDone_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDone.Click

        strError = String.Empty

        Dim item As Control
        Dim booValid As Boolean = True

        Dim owner_id As String
        Dim facility_id As String

        Dim cnt As Integer = 0

        If booValid = True Then

            For Each item In MyBase.Controls

                If InStr(item.GetType.ToString, "TextBox") Or InStr(LCase(item.GetType.ToString), "radiobutton") _
                    Or InStr(LCase(item.GetType.ToString), "datetime") _
                    Or InStr(LCase(item.GetType.ToString), "combobox") Then

                    Dim strValue As String = ""
                    If InStr(item.GetType.ToString, "TextBox") Then
                        strValue = item.Text
                        If strValue.Length = 0 Then
                            strValue = " "
                        End If
                    End If

                    If InStr(LCase(item.GetType.ToString), "radiobutton") Then
                        strValue = item.Text
                    End If

                    If InStr(LCase(item.GetType.ToString), "datetime") Then
                        Dim dtpicker As DateTimePicker = item
                        strValue = dtpicker.Value
                    End If

                    If InStr(LCase(item.GetType.ToString), "combobox") Then

                        Dim cmbValue As ComboBox = item

                        If cmbValue.SelectedValue Is Nothing OrElse TypeOf cmbValue.SelectedValue Is DBNull Then
                            strValue = cmbValue.Text
                        Else
                            strValue = cmbValue.SelectedValue
                        End If

                    End If


                    If CheckValue(item.Name, strValue, 4) = True Then
                        Dim strName As String = Microsoft.VisualBasic.Right(item.Name, Len(item.Name) - 4)

                        '
                        'RaiseEvent ReturnParameters(colParams)
                        ' Declare the parameter related objects.
                        '

                        Dim crParameterDiscreteValue As ParameterDiscreteValue
                        Dim crParameterFieldDefinitions As ParameterFieldDefinitions
                        Dim crParameterFieldLocation As ParameterFieldDefinition
                        Dim crParameterValues As ParameterValues
                        '
                        ' Get the report's parameters collection.

                        crParameterFieldDefinitions = Cr.DataDefinition.ParameterFields
                        '
                        ' Set the first parameter
                        ' - Get the parameter, tell it to use the current values vs default value.
                        ' - Tell it the parameter contains 1 discrete value vs multiple values.
                        ' - Set the parameter's value.
                        ' - Add it and apply it.
                        ' - Repeat these statements for each parameter.
                        '
                        crParameterValues = crParameterFieldDefinitions.Item(strName).CurrentValues

                        crParameterValues.AddValue(strValue)



                        Cr.DataDefinition.ParameterFields.Item(strName).ApplyCurrentValues(crParameterValues)



                        '    If subReportIndex > -1 And cnt = (subReportIndex - 1) Then

                        '   For g As Integer = subReportIndex To (Cr.DataDefinition.ParameterFields.Count - 1)

                        '  crParameterFieldLocation = Cr.DataDefinition.ParameterFields(g)

                        ' crParameterFieldLocation.CurrentValues.Clear()

                        'For Each cv As ParameterDiscreteValue In Cr.ParameterFields(subReportIndex - 1).CurrentValues


                        'Dim p As New ParameterDiscreteValue

                        'p.Value = cv.Value
                        'crParameterFieldLocation.CurrentValues.Add(p)

                        'Next

                        'crParameterFieldLocation.ApplyCurrentValues(crParameterFieldLocation.CurrentValues)
                        'Next g
                        ' End If
                    End If
                    cnt += 1

                End If

            Next


            If Cr.FileName.ToUpper.IndexOf("CAPCURRENTSUMMARYREPORT.RPT") > -1 AndAlso MsgBox("Would you like Update the Current CAP data within the database?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then

                Dim cap As CAP_Letters
                Try
                    cap = New CAP_Letters


                    cap.SetupSystemToGenerateCAPYearly(CAP_Letters.CapAnnualMode.CurrentSummary, False, Nothing, -1, Now.Year)
                Catch ex As Exception

                    Dim MyErr As New ErrorReport(ex)
                    MyErr.ShowDialog()
                Finally
                    cap = Nothing
                End Try
            End If



            RaiseEvent ReturnReport()

            Close()
        Else

            MsgBox(strError)
        End If

    End Sub


    Sub DisableCheckBox(ByVal sender As Object, ByVal e As EventArgs)

        With DirectCast(sender, CheckBox)

            Dim dp As DateTimePicker = DirectCast(.Tag, DateTimePicker)

            If .Checked Then

                dp.Enabled = False

                If .Name.ToUpper.IndexOf("BEGIN") > -1 Then
                    dp.Value = New Date(1988, 12, 22)
                Else
                    dp.Value = New Date(1900, 1, 1)
                End If

            Else

                dp.Enabled = True
                dp.Value = New Date(Now.Year, Now.Month, Now.Day)

            End If

        End With
    End Sub
#End Region
End Class

