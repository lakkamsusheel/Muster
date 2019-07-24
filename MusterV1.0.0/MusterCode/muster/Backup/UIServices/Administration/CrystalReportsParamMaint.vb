Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class CrystalReportsParamMaint
    Inherits System.Windows.Forms.Form
    Dim strError As String
    Public Cr As ReportDocument
    Public WithEvents ReportParams As Muster.BusinessLogic.pReportParams
    Friend Event ReturnReport()
    Dim returnVal As String = String.Empty

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

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
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnDone = New System.Windows.Forms.Button
        Me.lblMsg = New System.Windows.Forms.Label
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'btnDone
        '
        Me.btnDone.Location = New System.Drawing.Point(112, 72)
        Me.btnDone.Name = "btnDone"
        Me.btnDone.Size = New System.Drawing.Size(176, 24)
        Me.btnDone.TabIndex = 0
        Me.btnDone.Text = "Save descriptions"
        Me.btnDone.Visible = False
        '
        'lblMsg
        '
        Me.lblMsg.Location = New System.Drawing.Point(32, 8)
        Me.lblMsg.Name = "lblMsg"
        Me.lblMsg.Size = New System.Drawing.Size(344, 32)
        Me.lblMsg.TabIndex = 1
        Me.lblMsg.Text = "Please supply the user friendly descriptions for the following report parameters." & _
        ""
        '
        'btnSave
        '
        Me.btnSave.Enabled = False
        Me.btnSave.Location = New System.Drawing.Point(116, 72)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(64, 24)
        Me.btnSave.TabIndex = 2
        Me.btnSave.Text = "Save"
        '
        'btnCancel
        '
        Me.btnCancel.Enabled = False
        Me.btnCancel.Location = New System.Drawing.Point(180, 72)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(64, 24)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "Cancel"
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(244, 72)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(64, 24)
        Me.btnClose.TabIndex = 4
        Me.btnClose.Text = "Close"
        '
        'CrystalReportsParamMaint
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(416, 102)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.lblMsg)
        Me.Controls.Add(Me.btnDone)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "CrystalReportsParamMaint"
        Me.Text = "Report Parameters"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub CrystalReportsTestParamForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim i As Integer = 0
        Dim intTop As Integer = 0
        For i = 0 To Cr.DataDefinition.ParameterFields.Count - 1
            ReportParams.Retrieve("SYSTEM|REPORTPARAMS|" & ReportParams.ReportID & "|" & Cr.DataDefinition.ParameterFields(i).Name, False)

            'To skip the Sub-Report parameters.
            If Not Cr.DataDefinition.ParameterFields(i).ReportName = String.Empty Then
                Exit For
            End If

            Dim label As New Label
            label.Text = Cr.DataDefinition.ParameterFields(i).Name() '& "(undefined)"
            label.Left = 30
            label.Top = intTop + 50
            label.Width = 150
            Me.Controls.Add(label)

            Dim strName As String
            Dim strType As String
            strName = "str_" & Cr.DataDefinition.ParameterFields(i).Name
            strType = "str"

            'If strType = "str" Then
            Dim textbox As New TextBox
            textbox.Name = strName
            textbox.Left = 190
            textbox.Top = intTop + 50

            'Get the Prompting text when the parameter description is empty in the database.
            textbox.Text = IIf(ReportParams.ParamDescription = String.Empty, Cr.DataDefinition.ParameterFields(i).PromptText, ReportParams.ParamDescription)
            textbox.Tag = "SYSTEM|REPORTPARAMS|" & ReportParams.ReportID & "|" & Cr.DataDefinition.ParameterFields(i).Name
            Me.Controls.Add(textbox)

            AddHandler Me.Controls.Item(Me.Controls.GetChildIndex(textbox)).Leave, AddressOf TextBox_Leave_Check

            'ElseIf strType = "boo" Then
            '    Dim rdlist As New RadioButton
            '    rdlist.Name = strName
            '    rdlist.Text = "TRUE"
            '    rdlist.Checked = True
            '    rdlist.Left = 190
            '    rdlist.Top = intTop + 50
            '    Me.Controls.Add(rdlist)
            '    rdlist = New RadioButton
            '    rdlist.Name = strName
            '    rdlist.Text = "FALSE"
            '    rdlist.Checked = False
            '    rdlist.Left = 190
            '    rdlist.Top = intTop + 70
            '    Me.Controls.Add(rdlist)
            'ElseIf strType = "dat" Then
            '    Dim dtPicker As New DateTimePicker
            '    dtPicker.Name = strName
            '    dtPicker.Left = 190
            '    dtPicker.Top = intTop + 50
            '    Me.Controls.Add(dtPicker)
            'Else
            '    Dim textbox As New TextBox
            '    textbox.Name = strName
            '    textbox.Left = 190
            '    textbox.Top = intTop + 50
            '    Me.Controls.Add(textbox)
            'End If

            Cr.DataDefinition.ParameterFields.MoveNext()

            intTop = intTop + 50
        Next

        Me.btnDone.Top = intTop + 100
        Me.Height = intTop + 200
    End Sub

    Private Sub TextBox_Leave_Check(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim MyInfo As Muster.Info.ProfileInfo
        MyInfo = ReportParams.Retrieve(sender.tag)
        If MyInfo Is Nothing Then
            Dim MyArr() As String
            MyArr = sender.tag.split("|")
            MyInfo = New Muster.Info.ProfileInfo(MyArr(0), MyArr(1), MyArr(2), MyArr(3), "", False, MusterContainer.AppUser.ID, Now, MusterContainer.AppUser.ID, Now)
            ReportParams.Add(MyInfo)
            MyInfo = ReportParams.Retrieve(sender.tag)
        End If
        ReportParams.ParamDescription = sender.text
    End Sub
    Private Sub btnDone_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDone.Click
        strError = ""
        Dim item As Control
        Dim booValid As Boolean = True

        For Each item In MyBase.Controls
            If InStr(item.GetType.ToString, "TextBox") Or InStr(LCase(item.GetType.ToString), "radiobutton") Or InStr(LCase(item.GetType.ToString), "datetime") Then
                Dim strValue As String = ""
                If InStr(item.GetType.ToString, "TextBox") Then
                    strValue = item.Text
                End If
                If InStr(LCase(item.GetType.ToString), "radiobutton") Then
                    strValue = item.Text
                End If
                If InStr(LCase(item.GetType.ToString), "datetime") Then
                    Dim dtpicker As DateTimePicker = item
                    strValue = dtpicker.Value
                End If
                Dim strName As String = Microsoft.VisualBasic.Right(item.Name, Len(item.Name) - 4)
                Dim oParamInfo As New Muster.Info.ProfileInfo
                oParamInfo = ReportParams.Retrieve("SYSTEM|REPORTPARAMS|" & ReportParams.ReportID & "|" & strName, False)
                If oParamInfo Is Nothing Then
                    ReportParams.Add(New Muster.Info.ProfileInfo("SYSTEM", "REPORTPARAMS", ReportParams.ReportID, strName, "", False, "", Now(), "", Now()))
                End If

                ReportParams.ParamDescription = strValue
            End If
        Next
        Me.Close()
    End Sub
#Region "External Event Handlers"
    Private Sub ParamChanged(ByVal bolValue As Boolean) Handles ReportParams.ReportParamChanged
        btnSave.Enabled = bolValue
        btnCancel.Enabled = bolValue
    End Sub
#End Region

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)
        If ReportParams.colIsDirty Then
            Dim Results As Long = MsgBox("There are unsaved changes. Do you want to save changes before closing?", MsgBoxStyle.Question & MsgBoxStyle.YesNoCancel, "Parameter Description(s) Changed")
            If Results = MsgBoxResult.Yes Then
                ReportParams.Flush(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal)
            Else
                If Results = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
            End If
        End If
    End Sub

    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click

        ReportParams.Save(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal)
        MsgBox("Save Successful")
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim oTextBox As Windows.Forms.TextBox
        Dim oControl As Windows.Forms.Control
        ReportParams.Reset()
        For Each oControl In Me.Controls
            If oControl.GetType.FullName = "System.Windows.Forms.TextBox" Then
                If oControl.Tag = ReportParams.ProfileID Then
                    oControl.Text = ReportParams.ParamDescription
                End If
            End If
        Next
    End Sub
End Class

