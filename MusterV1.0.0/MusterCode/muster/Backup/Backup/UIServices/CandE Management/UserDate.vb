Public Class UserDate
    Inherits System.Windows.Forms.Form
    Private pLCE As MUSTER.BusinessLogic.pLicenseeComplianceEvent
    Private nPendingLetterTemplateNum As Integer = 0
    Private bolSave As Boolean = False
    Private bolCancel As Boolean = False
    Public bolEscalationCancelled As Boolean = False
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub
    Public Sub New(ByVal str As String, ByRef oLCE As MUSTER.BusinessLogic.pLicenseeComplianceEvent, ByVal pendingLetterTempNum As Integer)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        pLCE = oLCE
        lblDate.Text = str
        nPendingLetterTemplateNum = pendingLetterTempNum
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
    Friend WithEvents dtPickerDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents lblDate As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.dtPickerDate = New System.Windows.Forms.DateTimePicker
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.lblDate = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'dtPickerDate
        '
        Me.dtPickerDate.Checked = False
        Me.dtPickerDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickerDate.Location = New System.Drawing.Point(144, 8)
        Me.dtPickerDate.Name = "dtPickerDate"
        Me.dtPickerDate.ShowCheckBox = True
        Me.dtPickerDate.Size = New System.Drawing.Size(88, 20)
        Me.dtPickerDate.TabIndex = 0
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(56, 40)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.TabIndex = 1
        Me.btnSave.Text = "Save"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(136, 40)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.Text = "Cancel"
        '
        'lblDate
        '
        Me.lblDate.Location = New System.Drawing.Point(32, 8)
        Me.lblDate.Name = "lblDate"
        Me.lblDate.Size = New System.Drawing.Size(100, 25)
        Me.lblDate.TabIndex = 3
        Me.lblDate.Text = "Show Cause Hearing Date"
        '
        'UserDate
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(264, 70)
        Me.Controls.Add(Me.lblDate)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.dtPickerDate)
        Me.Name = "UserDate"
        Me.Text = "UserDate"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub UserDate_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        UIUtilsGen.CreateEmptyFormatDatePicker(dtPickerDate)
    End Sub

    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        bolSave = True
        If dtPickerDate.Text <> String.Empty And dtPickerDate.Text <> "__/__/____" Then
            If lblDate.Text.ToUpper = "SHOW CAUSE HEARING" Then
                pLCE.NextDueDate = pLCE.ShowCauseDate
                pLCE.PendingLetter = 1006
                pLCE.PendingLetterTemplateNum = nPendingLetterTemplateNum
                pLCE.PendingLetterName = "show cause hearing".ToUpper
                pLCE.OverrideDueDate = CDate("01/01/0001")
            ElseIf lblDate.Text.ToUpper = "COMMISSION HEARING" Then
                pLCE.NextDueDate = pLCE.CommissionDate
                pLCE.PendingLetter = 1008
                pLCE.PendingLetterTemplateNum = nPendingLetterTemplateNum
                pLCE.PendingLetterName = "Commission hearing".ToUpper
                pLCE.OverrideDueDate = CDate("01/01/0001")
            End If
            Me.Close()
        Else
            MsgBox(" Date is required")
        End If
    End Sub

    Private Sub dtPickerDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickerDate.ValueChanged
        UIUtilsGen.ToggleDateFormat(Me.dtPickerDate)
        If lblDate.Text.ToUpper = "SHOW CAUSE HEARING" Then
            UIUtilsGen.FillDateobjectValues(pLCE.ShowCauseDate, dtPickerDate.Text)
        ElseIf lblDate.Text.ToUpper = "COMMISSION HEARING" Then
            UIUtilsGen.FillDateobjectValues(pLCE.CommissionDate, dtPickerDate.Text)
        End If
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        bolCancel = True
        MsgBox(lblDate.Text + " cannot be null. Escalation process is cancelled")
        pLCE.Reset()
        bolEscalationCancelled = True
        Me.Close()
    End Sub

    Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)
        Dim sender As Object
        Dim en As System.EventArgs
        If Not bolSave And Not bolCancel Then
            bolSave = False
            bolCancel = False
            If dtPickerDate.Text <> String.Empty And dtPickerDate.Text <> "__/__/____" Then
                Dim Results As Long = MsgBox("There are unsaved changes.  Do you wish to save the changes?", MsgBoxStyle.Question & MsgBoxStyle.YesNoCancel, "Data Changed")
                If Results = MsgBoxResult.Yes Then
                    Me.btnSave_Click(sender, en)
                Else
                    Me.btnCancel_Click(sender, en)
                End If
            Else
                Me.btnCancel_Click(sender, en)
            End If
        End If
    End Sub
End Class

